import { spawn, type ChildProcess } from "node:child_process";
import { mkdtemp, readFile, rm, unlink } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { errorMessage, execFileAsync, findExecutable } from "./exec.js";

// Rasterizing HTML needs a browser engine, and the one that is already on almost
// every machine is Chrome. We shell out to it exactly as `render.ts` shells out
// to LibreOffice: detect the binary, and if it is absent degrade gracefully
// rather than failing the build.
//
// Set $CHROME_PATH to point at a specific binary.
const CHROME_CANDIDATES = [
  // macOS
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
  "/Applications/Chromium.app/Contents/MacOS/Chromium",
  "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
  "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
  // Linux / PATH
  "google-chrome",
  "google-chrome-stable",
  "chromium",
  "chromium-browser",
  "microsoft-edge",
  "/snap/bin/chromium",
  // Windows
  `${process.env.PROGRAMFILES ?? ""}\\Google\\Chrome\\Application\\chrome.exe`,
  `${process.env["PROGRAMFILES(X86)"] ?? ""}\\Google\\Chrome\\Application\\chrome.exe`,
  `${process.env.LOCALAPPDATA ?? ""}\\Google\\Chrome\\Application\\chrome.exe`,
  `${process.env.PROGRAMFILES ?? ""}\\Microsoft\\Edge\\Application\\msedge.exe`
];

const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
// A PNG always ends with a zero-length IEND chunk. Seeing it is how we know the
// file on disk is complete rather than half-written.
const PNG_IEND = Buffer.from([0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82]);

const POLL_INTERVAL_MS = 25;
const DEFAULT_TIMEOUT_MS = 30_000;

export type ShotRequest = {
  htmlPath: string;
  pngPath: string;
  width: number;
  height: number;
  scale: number;
  transparent?: boolean;
  allowNetwork?: boolean;
  timeoutMs?: number;
};

/** Rasterize a local HTML file. Throws with a readable message on failure. */
export type ShotFn = (request: ShotRequest) => Promise<{ pxWidth: number; pxHeight: number }>;

let chromePromise: Promise<string | undefined> | undefined;

/** Locate a Chrome-family browser. Memoized: detection runs at most once. */
export function findChrome(): Promise<string | undefined> {
  chromePromise ??= findExecutable(CHROME_CANDIDATES, "CHROME_PATH");
  return chromePromise;
}

const versionCache = new Map<string, string>();

/**
 * The browser's major version. It is part of the figure cache key, so that
 * upgrading Chrome re-renders figures instead of leaving stale PNGs forever.
 */
export async function chromeVersion(bin: string): Promise<string> {
  const cached = versionCache.get(bin);
  if (cached) return cached;
  let major = "unknown";
  try {
    const { stdout } = await execFileAsync(bin, ["--version"], { timeout: 10_000 });
    major = stdout.trim().match(/(\d+)\.\d+/)?.[1] ?? "unknown";
  } catch {
    // A browser that cannot report its version can still take screenshots.
  }
  versionCache.set(bin, major);
  return major;
}

export async function screenshotHtml(request: ShotRequest): Promise<{ pxWidth: number; pxHeight: number }> {
  const bin = await findChrome();
  if (!bin) throw new Error("No Chrome-family browser was found.");

  // The poll loop below treats "a complete PNG exists" as success, so a leftover
  // file from an earlier run would be mistaken for this run's output.
  await unlink(request.pngPath).catch(() => {});

  const profileDir = await mkdtemp(path.join(os.tmpdir(), "pptx-gen-chrome-"));
  try {
    const args = [
      "--headless=new",
      "--disable-gpu",
      "--hide-scrollbars",
      "--mute-audio",
      "--no-first-run",
      "--no-default-browser-check",
      "--disable-extensions",
      "--disable-background-networking",
      "--disable-component-update",
      "--disable-sync",
      "--force-color-profile=srgb",
      "--font-render-hinting=none",
      // Builds must not depend on the network. Blocking resolution at the
      // process boundary makes that true rather than merely intended: a figure
      // that reaches for a CDN font fails to load it here, exactly as it would
      // on a machine with no connection.
      ...(request.allowNetwork ? [] : ["--host-resolver-rules=MAP * ~NOTFOUND,EXCLUDE localhost"]),
      ...(process.platform === "linux" ? ["--no-sandbox", "--disable-dev-shm-usage"] : []),
      `--user-data-dir=${profileDir}`,
      `--window-size=${request.width},${request.height}`,
      `--force-device-scale-factor=${request.scale}`,
      `--default-background-color=${request.transparent ? "00000000" : "FFFFFFFF"}`,
      "--virtual-time-budget=1500",
      // Must be absolute: bare `--screenshot` writes into the process's cwd.
      `--screenshot=${path.resolve(request.pngPath)}`,
      pathToFileURL(path.resolve(request.htmlPath)).href
    ];

    const bytes = await runUntilPng(bin, args, request.pngPath, request.timeoutMs ?? DEFAULT_TIMEOUT_MS);
    return readPngSize(bytes);
  } finally {
    // Best effort only. The browser's children may still be flushing into the
    // profile as it is torn down, and a throwaway temp directory that outlives
    // us by a few seconds must never fail a screenshot that already succeeded.
    await rm(profileDir, { recursive: true, force: true }).catch(() => {});
  }
}

/**
 * Run the browser until a complete PNG appears on disk, then stop it.
 *
 * Headless Chrome is not a well-behaved batch tool here. Given its own
 * `--user-data-dir` it writes the screenshot and then keeps running instead of
 * exiting, and when it is killed for that it reports failure despite having
 * produced a perfect file. Waiting on the exit code would therefore mean waiting
 * for the timeout on every single figure, and then disbelieving the result.
 *
 * So the file is the contract: poll for a complete PNG, return as soon as one
 * lands, and treat the exit code purely as diagnostic text for the failure case.
 */
async function runUntilPng(bin: string, args: string[], pngPath: string, timeoutMs: number): Promise<Buffer> {
  const child = spawn(bin, args, { cwd: os.tmpdir(), stdio: ["ignore", "ignore", "pipe"] });
  let stderr = "";
  child.stderr?.on("data", (chunk) => {
    // Fontconfig, DBus and GPU chatter is normal on every platform; keep a
    // bounded tail only so a genuine failure has something to report.
    stderr = (stderr + String(chunk)).slice(-2000);
  });

  let exited = false;
  let spawnError: Error | undefined;
  child.on("exit", () => { exited = true; });
  child.on("error", (error) => { spawnError = error; exited = true; });

  try {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      await delay(POLL_INTERVAL_MS);
      const bytes = await readFile(pngPath).catch(() => undefined);
      if (bytes && isCompletePng(bytes)) return bytes;
      if (exited) {
        if (bytes && isCompletePng(bytes)) return bytes;
        break;
      }
    }
  } finally {
    if (!exited) {
      child.kill("SIGKILL");
      // Give it a moment to actually die, so the profile directory is quiet
      // before the caller removes it.
      await waitForExit(child, 2000);
    }
  }

  if (spawnError) throw new Error(`Could not start the browser: ${errorMessage(spawnError)}`);
  const detail = stderr.trim().split("\n").slice(-3).join(" ").trim();
  throw new Error(`The browser produced no screenshot${detail ? `: ${detail}` : "."}`);
}

function isCompletePng(bytes: Buffer): boolean {
  return bytes.length > 24
    && bytes.subarray(0, 8).equals(PNG_MAGIC)
    && bytes.subarray(bytes.length - 12).equals(PNG_IEND);
}

/**
 * Read a PNG's real pixel dimensions from its IHDR header.
 *
 * Always prefer this to the requested viewport: it is the rendered truth, so the
 * fitting maths downstream cannot be thrown off by a browser that interprets
 * `--window-size` or the device scale factor differently than expected.
 */
export function readPngSize(bytes: Buffer): { pxWidth: number; pxHeight: number } {
  if (bytes.length < 24 || !bytes.subarray(0, 8).equals(PNG_MAGIC)) {
    throw new Error("Not a PNG file.");
  }
  return { pxWidth: bytes.readUInt32BE(16), pxHeight: bytes.readUInt32BE(20) };
}

function waitForExit(child: ChildProcess, ms: number): Promise<void> {
  return new Promise((resolve) => {
    const timer = setTimeout(resolve, ms);
    timer.unref?.();
    child.once("exit", () => {
      clearTimeout(timer);
      resolve();
    });
  });
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
