import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises";
import { AssetNotFoundError, createAssetResolver } from "./assets.js";

type Dirs = { projectDir: string; assetsDir: string; bundledIconsDir: string };

async function makeDirs(t: { after: (fn: () => unknown) => void }): Promise<Dirs> {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptx-assets-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));

  const dirs = {
    projectDir: path.join(root, "workspace", "projects", "deck"),
    assetsDir: path.join(root, "workspace", "assets"),
    bundledIconsDir: path.join(root, "install", "assets", "icons")
  };
  await mkdir(dirs.projectDir, { recursive: true });
  await mkdir(path.join(dirs.assetsDir, "icons"), { recursive: true });
  await mkdir(dirs.bundledIconsDir, { recursive: true });
  return dirs;
}

test("a bare icon name falls back to the set bundled with the engine", async (t) => {
  const dirs = await makeDirs(t);
  await writeFile(path.join(dirs.bundledIconsDir, "rocket.svg"), "<svg/>", "utf8");

  const assets = createAssetResolver(dirs);

  assert.equal(assets.resolveIcon("rocket"), path.join(dirs.bundledIconsDir, "rocket.svg"));
  assert.equal(assets.resolveIcon("rocket.svg"), path.join(dirs.bundledIconsDir, "rocket.svg"));
});

test("a workspace icon shadows the bundled one of the same name", async (t) => {
  const dirs = await makeDirs(t);
  await writeFile(path.join(dirs.bundledIconsDir, "rocket.svg"), "<svg>bundled</svg>", "utf8");
  await writeFile(path.join(dirs.assetsDir, "icons", "rocket.svg"), "<svg>mine</svg>", "utf8");

  const assets = createAssetResolver(dirs);

  assert.equal(assets.resolveIcon("rocket"), path.join(dirs.assetsDir, "icons", "rocket.svg"));
});

test("an icon missing from both libraries names every directory searched", async (t) => {
  const dirs = await makeDirs(t);
  const assets = createAssetResolver(dirs);

  assert.throws(
    () => assets.resolveIcon("nope"),
    (error: Error) => {
      assert.ok(error instanceof AssetNotFoundError);
      assert.match(error.message, /Icon "nope" was not found/);
      assert.ok(error.message.includes(path.join(dirs.assetsDir, "icons")));
      assert.ok(error.message.includes(dirs.bundledIconsDir));
      return true;
    }
  );
});

test("an icon reference with a separator is a project file, not a library lookup", async (t) => {
  const dirs = await makeDirs(t);
  const assets = createAssetResolver(dirs);

  assert.equal(assets.resolveIcon("inputs/logo.svg"), path.join(dirs.projectDir, "inputs", "logo.svg"));
  assert.equal(assets.resolveIcon("/abs/logo.svg"), path.resolve("/abs/logo.svg"));
});

test("a logo comes from the workspace only, and is undefined when absent", async (t) => {
  const dirs = await makeDirs(t);

  // Present in the install but NOT in the workspace: must not be found, or a
  // client deck could end up carrying the tool's placeholder mark.
  await mkdir(path.join(dirs.bundledIconsDir, ".."), { recursive: true });
  await writeFile(path.join(dirs.bundledIconsDir, "..", "logo-mark-dark.png"), "x", "utf8");

  const assets = createAssetResolver(dirs);
  assert.equal(assets.resolveLogo("logo-mark-dark.png"), undefined);

  await writeFile(path.join(dirs.assetsDir, "logo-mark-dark.png"), "x", "utf8");
  assert.equal(assets.resolveLogo("logo-mark-dark.png"), path.join(dirs.assetsDir, "logo-mark-dark.png"));
});

test("project paths resolve against the deck folder", async (t) => {
  const dirs = await makeDirs(t);
  const assets = createAssetResolver(dirs);

  assert.equal(assets.resolveProjectPath("inputs/chart.png"), path.join(dirs.projectDir, "inputs", "chart.png"));
  assert.equal(assets.resolveProjectPath("/tmp/chart.png"), path.resolve("/tmp/chart.png"));
});

test("a workspace with no bundled set still resolves its own icons", async (t) => {
  const dirs = await makeDirs(t);
  await writeFile(path.join(dirs.assetsDir, "icons", "mine.svg"), "<svg/>", "utf8");

  const assets = createAssetResolver({ projectDir: dirs.projectDir, assetsDir: dirs.assetsDir });

  assert.equal(assets.resolveIcon("mine"), path.join(dirs.assetsDir, "icons", "mine.svg"));
  assert.deepEqual(assets.iconDirs, [path.join(dirs.assetsDir, "icons")]);
});
