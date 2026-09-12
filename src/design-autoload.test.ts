// The regression test for the whole design-as-data scheme.
//
// Project scripts read the design at MODULE TOP LEVEL — `custom.ts` files are
// documented to start with `const { LM, CW, LS } = LAYOUT;`, and they build
// lookup tables from `C.accent` the same way. That only works if the workspace
// design is already applied by the time the importing module's body runs.
//
// These tests run a real script in a child process against a real workspace,
// because that ordering cannot be observed from inside a single test process.
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { mkdir, mkdtemp, realpath, rm, writeFile } from "node:fs/promises";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const INSTALL = path.resolve(HERE, "..");
// The child runs from a temp dir outside the install, where a bare `tsx`
// specifier would not resolve, so hand it an absolute one.
const TSX_LOADER = import.meta.resolve("tsx");

/** A workspace whose design.yml changes a colour, a font, and a layout value. */
async function makeWorkspace(t: { after: (fn: () => unknown) => void }): Promise<string> {
  const dir = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-autoload-")));
  t.after(() => rm(dir, { recursive: true, force: true }));

  await writeFile(path.join(dir, "pptx-gen.config.yml"), "version: 1\n", "utf8");
  await writeFile(
    path.join(dir, "design.yml"),
    ["colors:", "  accent: E8452C", "fonts:", "  sans: Roboto", "layout:", "  LM: 1.5"].join("\n"),
    "utf8"
  );
  return dir;
}

/**
 * A script that reads the design at top level, exactly the way a project's
 * custom.ts does, then prints what it captured.
 */
const TOP_LEVEL_READER = `
import { C, FONTS, LAYOUT } from ${JSON.stringify(path.join(INSTALL, "src", "index.ts"))};

// Captured at module scope, before anything calls into the engine.
const { LM } = LAYOUT;
const TONES = { accent: C.accent };
const sans = FONTS.sans;

console.log(JSON.stringify({ LM, accent: TONES.accent, sans }));
`;

async function runScript(
  scriptPath: string,
  cwd: string,
  env: NodeJS.ProcessEnv = {}
): Promise<Record<string, unknown>> {
  const { stdout } = await execFileAsync(process.execPath, ["--import", TSX_LOADER, scriptPath], {
    cwd,
    env: { ...process.env, PPTX_GEN_NO_AUTOLOAD: "", PPTX_GEN_WORKSPACE: "", ...env }
  });
  return JSON.parse(stdout.trim());
}

test("a top-level read sees the workspace design, not the engine defaults", async (t) => {
  const workspace = await makeWorkspace(t);
  const script = path.join(workspace, "projects", "demo", "build.ts");
  await mkdir(path.dirname(script), { recursive: true });
  await writeFile(script, TOP_LEVEL_READER, "utf8");

  const result = await runScript(script, workspace);

  assert.equal(result.LM, 1.5, "a top-level `const { LM } = LAYOUT` must see the workspace value");
  assert.equal(result.accent, "E8452C");
  assert.equal(result.sans, "Roboto");
});

test("the workspace is found by walking up, from any cwd", async (t) => {
  const workspace = await makeWorkspace(t);
  const script = path.join(workspace, "projects", "deep", "nested", "build.ts");
  await mkdir(path.dirname(script), { recursive: true });
  await writeFile(script, TOP_LEVEL_READER, "utf8");

  // Run from a directory that is not under the workspace at all.
  const result = await runScript(script, os.tmpdir());

  assert.equal(result.LM, 1.5, "discovery must start from the entry script, not the cwd");
});

test("no workspace means the engine defaults, not a crash", async (t) => {
  const dir = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-no-ws-")));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const script = path.join(dir, "build.ts");
  await writeFile(script, TOP_LEVEL_READER, "utf8");

  // HOME is redirected so a real ~/.pptx-gen on this machine cannot be found.
  const result = await runScript(script, dir, { HOME: dir, USERPROFILE: dir });

  assert.equal(result.LM, 0.75);
  assert.equal(result.accent, "3B82F6");
  assert.equal(result.sans, "Inter");
});

test("PPTX_GEN_NO_AUTOLOAD=1 pins the engine defaults", async (t) => {
  const workspace = await makeWorkspace(t);
  const script = path.join(workspace, "build.ts");
  await writeFile(script, TOP_LEVEL_READER, "utf8");

  const result = await runScript(script, workspace, { PPTX_GEN_NO_AUTOLOAD: "1" });

  assert.equal(result.LM, 0.75, "the escape hatch the test suite relies on must pin the defaults");
  assert.equal(result.accent, "3B82F6");
});

test("a malformed design.yml fails loudly rather than shipping an off-brand deck", async (t) => {
  const workspace = await makeWorkspace(t);
  await writeFile(path.join(workspace, "design.yml"), "colors:\n  accent: reddish\n", "utf8");
  const script = path.join(workspace, "build.ts");
  await writeFile(script, TOP_LEVEL_READER, "utf8");

  await assert.rejects(
    () => runScript(script, workspace),
    (error: Error & { stderr?: string }) => {
      assert.match(error.stderr ?? "", /colors\.accent must be a 6-digit hex colour/);
      return true;
    }
  );
});
