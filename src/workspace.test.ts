import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import {
  WORKSPACE_ENV_VAR,
  WorkspaceNotFoundError,
  installDir,
  resolveWorkspaceSync,
  tryResolveWorkspaceSync
} from "./workspace.js";

// Every test runs with an env that has no PPTX_GEN_WORKSPACE unless it sets
// one, so the suite never reads the developer's real workspace.
const NO_ENV: NodeJS.ProcessEnv = {};

async function makeWorkspace(root: string, config = "version: 1\n"): Promise<string> {
  await mkdir(root, { recursive: true });
  await writeFile(path.join(root, "pptx-gen.config.yml"), config, "utf8");
  return root;
}

async function tempDir(t: { after: (fn: () => unknown) => void }): Promise<string> {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-ws-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  // macOS tmp is a symlink (/var -> /private/var); resolve so path comparisons hold.
  return await import("node:fs/promises").then(({ realpath }) => realpath(dir));
}

test("the explicit flag wins over the env var", async (t) => {
  const dir = await tempDir(t);
  const flagged = await makeWorkspace(path.join(dir, "flagged"));
  const fromEnv = await makeWorkspace(path.join(dir, "env"));

  const workspace = resolveWorkspaceSync({
    explicit: flagged,
    env: { [WORKSPACE_ENV_VAR]: fromEnv },
    cwd: dir,
    entryDir: dir
  });

  assert.equal(workspace.root, flagged);
  assert.equal(workspace.source, "flag");
});

test("the env var wins over discovery", async (t) => {
  const dir = await tempDir(t);
  const fromEnv = await makeWorkspace(path.join(dir, "env"));
  const discovered = await makeWorkspace(path.join(dir, "discovered"));

  const workspace = resolveWorkspaceSync({
    env: { [WORKSPACE_ENV_VAR]: fromEnv },
    cwd: discovered,
    entryDir: discovered
  });

  assert.equal(workspace.root, fromEnv);
  assert.equal(workspace.source, "env");
});

test("discovery walks up from a deeply nested entry script", async (t) => {
  const dir = await tempDir(t);
  const root = await makeWorkspace(path.join(dir, "decks"));
  const deep = path.join(root, "projects", "q4-review", "inputs");
  await mkdir(deep, { recursive: true });

  const workspace = resolveWorkspaceSync({ entryDir: deep, cwd: os.tmpdir(), env: NO_ENV });

  assert.equal(workspace.root, root);
  assert.equal(workspace.source, "discovered");
});

test("the entry script's directory is searched before the cwd", async (t) => {
  const dir = await tempDir(t);
  const entryWorkspace = await makeWorkspace(path.join(dir, "entry"));
  const cwdWorkspace = await makeWorkspace(path.join(dir, "cwd"));

  const workspace = resolveWorkspaceSync({ entryDir: entryWorkspace, cwd: cwdWorkspace, env: NO_ENV });

  assert.equal(workspace.root, entryWorkspace);
});

test("config paths resolve against the config file even from a different cwd", async (t) => {
  const dir = await tempDir(t);
  const root = await makeWorkspace(
    path.join(dir, "decks"),
    ["version: 1", "templates: ../shared-templates", "design: brand/design.yml", "customize: brand/rules.md"].join("\n")
  );

  const workspace = resolveWorkspaceSync({ explicit: root, cwd: os.tmpdir(), entryDir: os.tmpdir(), env: NO_ENV });

  assert.equal(workspace.templatesDir, path.join(dir, "shared-templates"));
  assert.equal(workspace.designPath, path.join(root, "brand", "design.yml"));
  assert.equal(workspace.projectsDir, path.join(root, "projects"));
  assert.equal(workspace.iconsDir, path.join(root, "assets", "icons"));
  assert.equal(workspace.customizePath, path.join(root, "brand", "rules.md"));
});

test("an explicit workspace that has no config is an error, never a fallback", async (t) => {
  const dir = await tempDir(t);
  await makeWorkspace(path.join(dir, "real"));
  const empty = path.join(dir, "empty");
  await mkdir(empty, { recursive: true });

  assert.throws(
    () =>
      resolveWorkspaceSync({
        explicit: empty,
        cwd: path.join(dir, "real"),
        entryDir: path.join(dir, "real"),
        env: NO_ENV
      }),
    (error: Error) => {
      assert.ok(error instanceof WorkspaceNotFoundError);
      assert.deepEqual(error.searched, [empty]);
      return true;
    }
  );
});

test("the not-found error lists where it looked and how to fix it", async (t) => {
  const dir = await tempDir(t);
  const start = path.join(dir, "a", "b");
  await mkdir(start, { recursive: true });

  assert.throws(
    () => resolveWorkspaceSync({ cwd: start, entryDir: start, env: NO_ENV }),
    (error: Error) => {
      assert.ok(error instanceof WorkspaceNotFoundError);
      assert.match(error.message, /No pptx-gen workspace found\./);
      assert.match(error.message, /pptx-gen init ~\/decks/);
      assert.match(error.message, /PPTX_GEN_WORKSPACE/);
      assert.ok(error.searched.includes(start), "should list the directory it started from");
      assert.ok(error.searched.includes(path.join(os.homedir(), ".pptx-gen")), "should list the home workspace");
      return true;
    }
  );
});

test("tryResolveWorkspaceSync returns undefined instead of throwing", async (t) => {
  const dir = await tempDir(t);
  const start = path.join(dir, "nowhere");
  await mkdir(start, { recursive: true });

  assert.equal(tryResolveWorkspaceSync({ cwd: start, entryDir: start, env: NO_ENV }), undefined);
});

test("tryResolveWorkspaceSync swallows a malformed config", async (t) => {
  const dir = await tempDir(t);
  const root = await makeWorkspace(path.join(dir, "broken"), "version: 99\n");

  assert.equal(tryResolveWorkspaceSync({ explicit: root, env: NO_ENV }), undefined);
  assert.throws(() => resolveWorkspaceSync({ explicit: root, env: NO_ENV }), /unsupported version 99/);
});

test("the install dir is the engine root, and is never the workspace", async (t) => {
  const dir = await tempDir(t);
  const root = await makeWorkspace(path.join(dir, "decks"));
  const workspace = resolveWorkspaceSync({ explicit: root, env: NO_ENV });

  assert.equal(workspace.installDir, installDir());
  assert.notEqual(workspace.installDir, workspace.root);
  assert.equal(workspace.bundledIconsDir, path.join(installDir(), "assets", "icons"));
});
