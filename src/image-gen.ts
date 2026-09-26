// Generating the images a brief asks for — the one part of pptx-gen that talks
// to a network service, and the reason it is a command of its own.
//
// `pptx-gen images` reads the brief's `## Images` block and customize.md's
// `## Image generation` section, composes a prompt per image, asks the image
// model for the variants, and writes them into the deck's `inputs/` folder.
// The build then places them like any other input. It never runs during a
// build: build scripts stay offline and deterministic (docs/decisions/0010).
//
// The feature is optional. With no OPENAI_API_KEY it generates nothing, says
// so in the brief, and the build draws captioned placeholders instead.
import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import util from "node:util";
import { C } from "./design.js";
import {
  brandPalette,
  composeImagePrompt,
  type ImageRequest,
  type ImageStyle,
  NO_IMAGE_STYLE,
  parseImageRequests,
  parseImageStyle,
  sizeFor,
  updateImageNotes
} from "./image-brief.js";
import { BRIEF_FILE, INPUTS_DIR, VARIANT_EXTENSIONS } from "./images.js";
import type { Workspace } from "./workspace.js";
import { closestMatch } from "./workspace-config.js";

export type ImageFormat = "jpeg" | "png";

type ImageReference = { name: string; bytes: Buffer; mime: string };

export type ImageGenRequest = {
  model: string;
  prompt: string;
  /** `WxH`. */
  size: string;
  quality: string;
  format: ImageFormat;
  /** How many images to return. */
  n: number;
  /** Example images to match. Empty for a plain text-to-image request. */
  references: ImageReference[];
};

/**
 * Turns one request into image bytes. The seam tests inject, the way
 * `ShotFn` stands in for Chrome, so no test ever reaches the network.
 */
export type ImageGenFn = (request: ImageGenRequest) => Promise<Buffer[]>;

export const DEFAULT_IMAGE_MODEL = "gpt-image-2.5-flare";
const DEFAULT_IMAGE_QUALITY = "medium";
const DEFAULT_IMAGE_FORMAT: ImageFormat = "jpeg";
export const LOCK_FILE = "images.lock.json";

const DEFAULT_BASE_URL = "https://api.openai.com/v1";
const MAX_REFERENCES = 16;
const REFERENCE_TYPES: Record<string, string> = {
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".webp": "image/webp"
};

/** Whether image generation is on here, and with what settings. */
export type ImageGeneration = {
  available: boolean;
  /** Where the key came from. The key itself is never printed. */
  keySource: "env" | ".env" | null;
  model: string;
  quality: string;
  baseUrl: string;
  apiKey?: string;
};

// `util.parseEnv` arrived in Node 20.12. Reached through the default import so
// an older Node still loads this module — a named import would fail to link,
// taking every CLI command down with it — and simply skips the .env file.
const parseEnv: ((content: string) => NodeJS.Dict<string>) | undefined = util.parseEnv;

/**
 * Look for an API key and settings: the environment first, then a `.env` file
 * in the workspace root. An empty value counts as unset. `process.env` is
 * never modified.
 */
export function resolveImageGeneration(
  workspaceRoot: string | undefined,
  env: NodeJS.ProcessEnv = process.env
): ImageGeneration {
  const fileEnv = workspaceRoot ? readDotEnv(path.join(workspaceRoot, ".env")) : {};
  const read = (name: string) => nonEmpty(env[name]) ?? nonEmpty(fileEnv[name]);

  const fromEnv = nonEmpty(env.OPENAI_API_KEY);
  const apiKey = fromEnv ?? nonEmpty(fileEnv.OPENAI_API_KEY);
  return {
    available: apiKey !== undefined,
    keySource: fromEnv ? "env" : apiKey ? ".env" : null,
    model: read("PPTX_GEN_IMAGE_MODEL") ?? DEFAULT_IMAGE_MODEL,
    quality: read("PPTX_GEN_IMAGE_QUALITY") ?? DEFAULT_IMAGE_QUALITY,
    baseUrl: read("OPENAI_BASE_URL") ?? DEFAULT_BASE_URL,
    apiKey
  };
}

function readDotEnv(file: string): NodeJS.Dict<string> {
  if (!parseEnv || !existsSync(file)) return {};
  try {
    return parseEnv(readFileSync(file, "utf8"));
  } catch {
    return {};
  }
}

function nonEmpty(value: string | undefined): string | undefined {
  const trimmed = value?.trim();
  return trimmed ? trimmed : undefined;
}

export type OpenAiImageGenOptions = {
  apiKey: string;
  baseUrl?: string;
  /** Injected by tests. */
  fetch?: typeof fetch;
  /** Injected by tests, so a retry does not really wait. */
  sleep?: (ms: number) => Promise<void>;
  timeoutMs?: number;
  maxAttempts?: number;
};

/**
 * The OpenAI Images API as an `ImageGenFn`.
 *
 * Plain `fetch`, no SDK. A request with references goes to `/images/edits` as
 * multipart, one `image[]` part per reference; one without goes to
 * `/images/generations` as JSON. Rate limits (429) and server errors are
 * retried with backoff, honouring `retry-after`; anything else fails at once
 * with the API's own message, because retrying a rejected prompt cannot help.
 */
export function openAiImageGen(options: OpenAiImageGenOptions): ImageGenFn {
  const fetchFn = options.fetch ?? globalThis.fetch;
  const sleep = options.sleep ?? ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  const baseUrl = (options.baseUrl ?? DEFAULT_BASE_URL).replace(/\/+$/, "");
  const timeoutMs = options.timeoutMs ?? 180_000;
  const maxAttempts = options.maxAttempts ?? 4;

  return async (request) => {
    const withReferences = request.references.length > 0;
    const url = `${baseUrl}/images/${withReferences ? "edits" : "generations"}`;

    for (let attempt = 1; ; attempt += 1) {
      let response: Response;
      try {
        response = await fetchFn(url, {
          method: "POST",
          headers: {
            Authorization: `Bearer ${options.apiKey}`,
            ...(withReferences ? {} : { "Content-Type": "application/json" })
          },
          body: withReferences ? multipartBody(request) : jsonBody(request),
          signal: AbortSignal.timeout(timeoutMs)
        });
      } catch (error) {
        if ((error as Error).name === "TimeoutError") {
          throw new Error(`OpenAI did not answer within ${Math.round(timeoutMs / 1000)}s.`);
        }
        if (attempt >= maxAttempts) throw new Error(`Could not reach OpenAI: ${(error as Error).message}`);
        await sleep(backoff(attempt));
        continue;
      }

      if (response.ok) return decodeImages(await response.json());

      const problem = await readProblem(response);
      const retryable = (response.status === 429 && problem.code !== "insufficient_quota") || response.status >= 500;
      if (retryable && attempt < maxAttempts) {
        await sleep(retryDelay(response, attempt));
        continue;
      }
      throw new ImageGenError(response.status, describeProblem(response.status, problem, request));
    }
  };
}

function jsonBody(request: ImageGenRequest): string {
  return JSON.stringify({
    model: request.model,
    prompt: request.prompt,
    n: request.n,
    size: request.size,
    quality: request.quality,
    output_format: request.format
  });
}

function multipartBody(request: ImageGenRequest): FormData {
  const form = new FormData();
  form.append("model", request.model);
  form.append("prompt", request.prompt);
  form.append("n", String(request.n));
  form.append("size", request.size);
  form.append("quality", request.quality);
  form.append("output_format", request.format);
  for (const reference of request.references) {
    form.append("image[]", new Blob([new Uint8Array(reference.bytes)], { type: reference.mime }), reference.name);
  }
  return form;
}

function decodeImages(body: unknown): Buffer[] {
  const data = typeof body === "object" && body !== null && "data" in body ? body.data : undefined;
  const images = Array.isArray(data)
    ? data.flatMap((item: unknown) =>
        typeof item === "object" && item !== null && "b64_json" in item && typeof item.b64_json === "string"
          ? [Buffer.from(item.b64_json, "base64")]
          : []
      )
    : [];
  if (images.length === 0) throw new Error("OpenAI answered, but with no images in the response.");
  return images;
}

/** A refusal from the image API, with its HTTP status. */
class ImageGenError extends Error {
  constructor(
    readonly status: number,
    message: string
  ) {
    super(message);
    this.name = "ImageGenError";
  }
}

type Problem = { message: string; code?: string };

async function readProblem(response: Response): Promise<Problem> {
  const text = await response.text().catch(() => "");
  try {
    const body: unknown = JSON.parse(text);
    const error = typeof body === "object" && body !== null && "error" in body ? body.error : undefined;
    if (typeof error === "object" && error !== null) {
      const message = "message" in error && typeof error.message === "string" ? error.message : "";
      const code = "code" in error && typeof error.code === "string" ? error.code : undefined;
      return { message: message || text || response.statusText, code };
    }
  } catch {
    // Not JSON; fall through to the raw text.
  }
  return { message: text.trim() || response.statusText };
}

function describeProblem(status: number, problem: Problem, request: ImageGenRequest): string {
  const code = problem.code ? ` (${problem.code})` : "";
  let hint = "";
  if (status === 401) hint = " Check OPENAI_API_KEY.";
  else if (status === 404 && /model/i.test(problem.message)) {
    hint = ` Pick another model with --model or PPTX_GEN_IMAGE_MODEL.`;
  } else if (status === 400 && /size/i.test(problem.message)) {
    hint = ` ${request.model} may not accept ${request.size}; try aspect landscape, portrait or square.`;
  }
  // OpenAI quotes a masked piece of a rejected key back. This message ends up
  // in the brief, which is often committed, so no piece of any key survives.
  const message = problem.message.replace(/\bsk-[^\s.,;]+/g, "sk-…");
  return `OpenAI returned ${status}${code}: ${message}${hint}`;
}

function retryDelay(response: Response, attempt: number): number {
  const ms = Number(response.headers.get("retry-after-ms"));
  if (Number.isFinite(ms) && ms > 0) return Math.min(ms, 60_000);
  const seconds = Number(response.headers.get("retry-after"));
  if (Number.isFinite(seconds) && seconds > 0) return Math.min(seconds * 1000, 60_000);
  return backoff(attempt);
}

function backoff(attempt: number): number {
  return 2000 * 2 ** (attempt - 1);
}

/** What happened to one image in a run. */
export type ImageOutcome = {
  id: string;
  /**
   * `supplied` means a file someone put in `inputs/` by hand was found and left
   * alone; `planned` is a dry run.
   */
  status: "generated" | "cached" | "supplied" | "failed" | "planned";
  /** Files written or kept, relative to the deck folder. */
  files: string[];
  prompt: string;
  size: string;
  /** The references actually sent. */
  references: string[];
  reason?: string;
  /** Problems that did not stop the image, such as a missing reference. */
  notes: string[];
};

export type GenerateImagesSummary = {
  projectDir: string;
  briefPath: string;
  available: boolean;
  keySource: ImageGeneration["keySource"];
  model: string;
  quality: string;
  format: ImageFormat;
  dryRun: boolean;
  images: ImageOutcome[];
  /** Whether the notes block in the brief changed. */
  notesUpdated: boolean;
};

export type GenerateImagesOptions = {
  /** The deck folder, holding `brief.md` and `inputs/`. */
  projectDir: string;
  /** Supplies customize.md and the `.env` fallback. */
  workspace?: Workspace;
  env?: NodeJS.ProcessEnv;
  /** Replaces the OpenAI call. When given, no API key is needed. */
  gen?: ImageGenFn;
  /** Limit the run to these image ids. */
  only?: string[];
  /** Regenerate even when cached, and replace hand-supplied files. */
  force?: boolean;
  /** Compose the prompts and report what would happen, without calling anything or writing any file. */
  dryRun?: boolean;
  model?: string;
  quality?: string;
  format?: ImageFormat;
  /** Brand colours for the palette hint. Defaults to the live design. */
  colors?: Record<string, string>;
  /** Wraps each generation, for progress output. */
  step?: <T>(label: string, fn: () => Promise<T>) => Promise<T>;
};

type LockEntry = {
  hash: string;
  model: string;
  quality: string;
  size: string;
  format: ImageFormat;
  variants: number;
  prompt: string;
  references: string[];
  files: string[];
};

type LockFile = { version: 1; images: Record<string, LockEntry> };

/**
 * Generate every image the brief declares, or the ones named in `only`.
 *
 * Each image is generated once and then cached: the lock file records a hash
 * of everything that shapes the result (prompt, model, quality, size, format,
 * variant count, reference bytes), and a matching hash with all files present
 * means nothing to do. Changing only `pick` never regenerates. A file at a
 * variant's path that the lock does not own was put there by a person, and is
 * never overwritten without `force`.
 *
 * Failures never throw: they are returned, and written as notes into the
 * brief so whoever reads it next sees why a slide has a placeholder. Only a
 * malformed brief or customize.md — an authoring error — throws.
 */
export async function generateImages(options: GenerateImagesOptions): Promise<GenerateImagesSummary> {
  const projectDir = path.resolve(options.projectDir);
  const briefPath = path.join(projectDir, BRIEF_FILE);
  if (!existsSync(briefPath)) {
    throw new Error(`No ${BRIEF_FILE} in ${projectDir}. Images are declared in its ## Images block.`);
  }
  const brief = await readFile(briefPath, "utf8");
  const requests = selectRequests(parseImageRequests(brief, briefPath), options.only);
  const style = readStyle(options.workspace);

  const generation = resolveImageGeneration(options.workspace?.root, options.env ?? process.env);
  const model = options.model ?? generation.model;
  const quality = options.quality ?? generation.quality;
  const format = options.format ?? DEFAULT_IMAGE_FORMAT;
  const gen =
    options.gen ??
    (generation.apiKey ? openAiImageGen({ apiKey: generation.apiKey, baseUrl: generation.baseUrl }) : undefined);
  const palette = brandPalette(options.colors ?? C);
  const step = options.step ?? (<T>(_label: string, fn: () => Promise<T>) => fn());
  const dryRun = options.dryRun === true;

  const lockPath = path.join(projectDir, INPUTS_DIR, LOCK_FILE);
  const lock = readLock(lockPath);
  const lockBefore = JSON.stringify(lock);

  const outcomes: ImageOutcome[] = [];
  // Set once the API rejects the key: every later call would fail the same way.
  let rejectedKey: string | undefined;
  for (const request of requests) {
    const references = await loadReferences(request, style, projectDir, options.workspace?.root);
    const prompt = composeImagePrompt(request, style, { palette, hasReferences: references.loaded.length > 0 });
    const size = sizeFor(request);
    const extension = format === "png" ? "png" : "jpg";
    const targets = Array.from({ length: request.variants }, (_unused, index) =>
      path.posix.join(INPUTS_DIR, `${request.id}-${index + 1}.${extension}`)
    );
    const hash = createHash("sha256")
      .update(
        JSON.stringify({
          prompt,
          model,
          quality,
          size,
          format,
          variants: request.variants,
          references: references.loaded.map((reference) => sha1(reference.bytes))
        })
      )
      .digest("hex");
    const base = {
      id: request.id,
      prompt,
      size,
      references: references.labels,
      notes: references.problems
    };
    const previous = lock.images[request.id];
    const owned = new Set(previous?.files ?? []);
    const exists = (file: string) => existsSync(path.join(projectDir, file));

    if (dryRun) {
      outcomes.push({ ...base, status: "planned", files: targets });
      continue;
    }
    if (!options.force && previous?.hash === hash && targets.every(exists)) {
      outcomes.push({ ...base, status: "cached", files: targets });
      continue;
    }

    const variantPaths = targets.flatMap((_target, index) =>
      VARIANT_EXTENSIONS.map((ext) => path.posix.join(INPUTS_DIR, `${request.id}-${index + 1}.${ext}`))
    );
    const supplied = variantPaths.filter((file) => exists(file) && !owned.has(file));
    if (!options.force && supplied.length > 0) {
      outcomes.push({
        ...base,
        status: "supplied",
        files: supplied,
        reason: `${supplied.join(", ")} was not generated here, so it was left alone (use --force to replace it).`
      });
      continue;
    }

    if (!gen) {
      outcomes.push({ ...base, status: "failed", files: [], reason: "OPENAI_API_KEY is not set." });
      continue;
    }
    if (rejectedKey) {
      outcomes.push({ ...base, status: "failed", files: [], reason: rejectedKey });
      continue;
    }

    try {
      const images = await step(`${request.id}  ${request.variants} variant(s), ${size}`, () =>
        gen({ model, prompt, size, quality, format, n: request.variants, references: references.loaded })
      );
      await mkdir(path.join(projectDir, INPUTS_DIR), { recursive: true });
      const written = targets.slice(0, images.length);
      for (const [index, file] of written.entries()) {
        await writeFile(path.join(projectDir, file), images[index]);
      }
      // Anything this image used to own, or (under --force) a same-numbered
      // file with another extension that would shadow the new one, goes.
      const stale = new Set([...owned, ...(options.force ? variantPaths : [])]);
      for (const file of stale) {
        if (!written.includes(file)) await rm(path.join(projectDir, file), { force: true });
      }
      lock.images[request.id] = {
        hash: written.length === targets.length ? hash : "",
        model,
        quality,
        size,
        format,
        variants: request.variants,
        prompt,
        references: references.labels,
        files: written
      };
      const notes = [...references.problems];
      if (written.length < targets.length) {
        notes.push(`only ${written.length} of ${targets.length} variants came back; rerun to fill in the rest.`);
      }
      outcomes.push({ ...base, status: "generated", files: written, notes });
    } catch (error) {
      if (error instanceof ImageGenError && error.status === 401) rejectedKey = error.message;
      outcomes.push({ ...base, status: "failed", files: [], reason: (error as Error).message });
    }
  }

  let notesUpdated = false;
  if (!dryRun) {
    if (JSON.stringify(lock) !== lockBefore) await writeLock(lockPath, lock);
    const next = updateImageNotes(brief, notesFor(outcomes, gen !== undefined));
    if (next !== brief) {
      await writeFile(briefPath, next, "utf8");
      notesUpdated = true;
    }
  }

  return {
    projectDir,
    briefPath,
    available: gen !== undefined,
    keySource: generation.keySource,
    model,
    quality,
    format,
    dryRun,
    images: outcomes,
    notesUpdated
  };
}

function selectRequests(requests: ImageRequest[], only: string[] | undefined): ImageRequest[] {
  if (!only || only.length === 0) return requests;
  const ids = requests.map((request) => request.id);
  for (const id of only) {
    if (ids.includes(id)) continue;
    const suggestion = closestMatch(id, ids);
    throw new Error(
      `No image '${id}' in the brief's ## Images block${suggestion ? ` — did you mean '${suggestion}'?` : ""}.`
    );
  }
  return requests.filter((request) => only.includes(request.id));
}

function readStyle(workspace: Workspace | undefined): ImageStyle {
  if (!workspace || !existsSync(workspace.customizePath)) return { ...NO_IMAGE_STYLE };
  return parseImageStyle(readFileSync(workspace.customizePath, "utf8"), workspace.root);
}

async function loadReferences(
  request: ImageRequest,
  style: ImageStyle,
  projectDir: string,
  workspaceRoot: string | undefined
): Promise<{ loaded: ImageReference[]; labels: string[]; problems: string[] }> {
  const candidates = [
    ...style.examples.map((file) => ({ file, label: labelFor(file, workspaceRoot) })),
    ...request.references.map((reference) => ({ file: resolveReference(reference, projectDir), label: reference }))
  ];
  const loaded: ImageReference[] = [];
  const labels: string[] = [];
  const problems: string[] = [];

  for (const { file, label } of candidates) {
    const mime = REFERENCE_TYPES[path.extname(file).toLowerCase()];
    if (!mime) {
      problems.push(`reference ${label} is not a PNG, JPEG or WebP, so it was not sent.`);
      continue;
    }
    if (!existsSync(file)) {
      problems.push(`reference ${label} was not found, so the image was made without it.`);
      continue;
    }
    if (loaded.length >= MAX_REFERENCES) {
      problems.push(`reference ${label} was not sent: at most ${MAX_REFERENCES} references go with one image.`);
      continue;
    }
    loaded.push({ name: path.basename(file), bytes: await readFile(file), mime });
    labels.push(label);
  }
  return { loaded, labels, problems };
}

function resolveReference(reference: string, projectDir: string): string {
  if (reference.startsWith("~/")) return path.join(os.homedir(), reference.slice(2));
  return path.resolve(projectDir, reference);
}

function labelFor(file: string, root: string | undefined): string {
  if (!root) return file;
  const relative = path.relative(root, file);
  return relative && !relative.startsWith("..") && !path.isAbsolute(relative) ? relative : file;
}

function notesFor(
  outcomes: ImageOutcome[],
  canGenerate: boolean
): { processed: string[]; notes: Map<string, string>; general?: string } {
  const notes = new Map<string, string>();
  const needsKey: string[] = [];

  for (const outcome of outcomes) {
    const parts: string[] = [];
    if (outcome.status === "failed") {
      if (!canGenerate) needsKey.push(outcome.id);
      else parts.push(`generation failed — ${outcome.reason} Rerun \`pptx-gen images\` to retry.`);
    }
    parts.push(...outcome.notes.map((note) => note.charAt(0).toUpperCase() + note.slice(1)));
    if (parts.length > 0) notes.set(outcome.id, parts.join(" "));
  }

  const general =
    needsKey.length > 0
      ? `OPENAI_API_KEY is not set, so ${needsKey.map((id) => `\`${id}\``).join(", ")} ${needsKey.length === 1 ? "was" : "were"} not generated. Set it in the environment or in the workspace's .env, then rerun \`pptx-gen images\`.`
      : undefined;
  return { processed: outcomes.map((outcome) => outcome.id), notes, general };
}

function readLock(lockPath: string): LockFile {
  const empty: LockFile = { version: 1, images: {} };
  if (!existsSync(lockPath)) return empty;
  let parsed: unknown;
  try {
    parsed = JSON.parse(readFileSync(lockPath, "utf8"));
  } catch {
    return empty;
  }
  const images = typeof parsed === "object" && parsed !== null && "images" in parsed ? parsed.images : undefined;
  if (typeof images !== "object" || images === null) return empty;

  const lock: LockFile = { version: 1, images: {} };
  for (const [id, entry] of Object.entries(images)) {
    if (
      typeof entry === "object" &&
      entry !== null &&
      "hash" in entry &&
      typeof entry.hash === "string" &&
      "files" in entry &&
      Array.isArray(entry.files) &&
      entry.files.every((file: unknown) => typeof file === "string")
    ) {
      lock.images[id] = entry as LockEntry;
    }
  }
  return lock;
}

async function writeLock(lockPath: string, lock: LockFile): Promise<void> {
  const images = Object.fromEntries(Object.entries(lock.images).sort(([a], [b]) => a.localeCompare(b)));
  await mkdir(path.dirname(lockPath), { recursive: true });
  await writeFile(lockPath, `${JSON.stringify({ version: 1, images }, null, 2)}\n`, "utf8");
}

function sha1(bytes: Buffer): string {
  return createHash("sha1").update(bytes).digest("hex");
}
