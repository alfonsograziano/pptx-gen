// End-to-end checks of the commands that create and repair a workspace.
//
// These shell out to the real CLI rather than calling the functions, because
// what matters is the whole path a user takes: init a folder somewhere else on
// disk, scaffold a deck in it, and build that deck with the engine resolved
// through the workspace's own link.
import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { execFile } from "node:child_process";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const INSTALL = path.resolve(HERE, "..");
const CLI = path.join(INSTALL, "src", "cli.ts");
const TSX_LOADER = import.meta.resolve("tsx");

async function cli(args: string[], cwd = INSTALL): Promise<{ stdout: string; stderr: string }> {
  return execFileAsync(process.execPath, ["--import", TSX_LOADER, CLI, ...args], {
    cwd,
    // Cleared so the suite's own pinned workspace cannot leak into these runs.
    env: { ...process.env, PPTX_GEN_WORKSPACE: "", PPTX_GEN_NO_AUTOLOAD: "" }
  });
}

async function tempDir(t: { after: (fn: () => unknown) => void }): Promise<string> {
  const dir = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-cli-test-")));
  t.after(() => rm(dir, { recursive: true, force: true }));
  return dir;
}

test("init creates a complete workspace outside the install", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);

  for (const file of [
    "pptx-gen.config.yml",
    "design.yml",
    "design.md",
    "customize.md",
    "package.json",
    "tsconfig.json",
    ".gitignore"
  ]) {
    assert.ok(existsSync(path.join(root, file)), `init should create ${file}`);
  }
  for (const dir of ["templates", "projects", "assets"]) {
    assert.ok(existsSync(path.join(root, dir)), `init should create ${dir}/`);
  }

  // Starter templates are copied, so the user owns and can delete them.
  assert.ok(existsSync(path.join(root, "templates", "title-cover", "template.pptx")));

  // The engine link is what makes `import "pptx-gen"` work from a project.
  assert.equal(await realpath(path.join(root, "node_modules", "pptx-gen")), INSTALL);

  // Nothing about the workspace lives inside the install.
  assert.ok(!root.startsWith(INSTALL));
});

test("init refuses to clobber an existing workspace, and --force is idempotent", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);
  await writeFile(path.join(root, "design.yml"), "colors:\n  accent: E8452C\n", "utf8");

  await assert.rejects(
    () => cli(["init", root]),
    (error: Error & { stderr?: string }) => {
      assert.match(error.stderr ?? "", /already exists — this is already a workspace/);
      return true;
    }
  );

  await cli(["init", root, "--force"]);
  assert.match(
    await readFile(path.join(root, "design.yml"), "utf8"),
    /E8452C/,
    "--force must not overwrite a brand the user has edited"
  );
});

test("init --no-starter-templates and --no-link leave those out", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root, "--no-starter-templates", "--no-link"]);

  assert.ok(existsSync(path.join(root, "templates")), "the directory is still created");
  assert.ok(!existsSync(path.join(root, "templates", "title-cover")));
  assert.ok(!existsSync(path.join(root, "node_modules", "pptx-gen")));
});

test("workspace --json reports absolute paths and the engine specifier", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);

  const { stdout } = await cli(["workspace", "--json", "--workspace", root]);
  const info = JSON.parse(stdout);

  assert.equal(info.root, root);
  assert.equal(info.templates, path.join(root, "templates"));
  assert.equal(info.projects, path.join(root, "projects"));
  assert.equal(info.design, path.join(root, "design.yml"));
  assert.equal(info.customize, path.join(root, "customize.md"));
  assert.equal(info.install, INSTALL);
  assert.equal(info.bundledIcons, path.join(INSTALL, "assets", "icons"));
  assert.equal(info.engineSpecifier, "pptx-gen");
  assert.equal(info.ok, true);
  assert.deepEqual(info.problems, []);

  for (const [key, value] of Object.entries(info)) {
    if (typeof value === "string" && key !== "engineSpecifier" && key !== "source") {
      assert.ok(path.isAbsolute(value), `${key} must be absolute, got ${value}`);
    }
  }
});

test("doctor --fix repairs a broken engine link", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);
  await rm(path.join(root, "node_modules"), { recursive: true, force: true });

  await assert.rejects(() => cli(["doctor", "--workspace", root]), /exit code 1|Command failed/);

  const { stdout } = await cli(["doctor", "--fix", "--workspace", root]);
  assert.match(stdout, /Linked/);
  assert.equal(await realpath(path.join(root, "node_modules", "pptx-gen")), INSTALL);

  const after = await cli(["doctor", "--workspace", root]);
  assert.match(after.stdout, /Workspace is healthy/);
});

test("doctor --fix restores a deleted customize.md", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);
  await rm(path.join(root, "customize.md"), { force: true });

  await assert.rejects(() => cli(["doctor", "--workspace", root]), /exit code 1|Command failed/);

  await cli(["doctor", "--fix", "--workspace", root]);
  assert.match(await readFile(path.join(root, "customize.md"), "utf8"), /# Customizations/);
});

test("a command outside any workspace explains how to make one", async (t) => {
  const dir = await tempDir(t);

  await assert.rejects(
    () => cli(["workspace"], dir),
    (error: Error & { stderr?: string }) => {
      assert.match(error.stderr ?? "", /No pptx-gen workspace found\./);
      assert.match(error.stderr ?? "", /pptx-gen init ~\/decks/);
      return true;
    }
  );
});

test("new scaffolds a deck that builds, with the engine imported by name", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);
  await cli(["new", "q4-review", "--custom", "--workspace", root]);

  const projectDir = path.join(root, "projects", "q4-review");
  const buildScript = path.join(projectDir, "build.ts");
  assert.ok(existsSync(buildScript));
  assert.match(await readFile(buildScript, "utf8"), /from "pptx-gen"/);
  assert.match(await readFile(path.join(projectDir, "custom.ts"), "utf8"), /from "pptx-gen"/);

  // Built from the install's cwd, with only the script path to go on: the
  // workspace has to be discovered from the script's own location.
  await cli(["build", "--script", buildScript]);

  assert.ok(existsSync(path.join(projectDir, "output", "deck.pptx")));
  assert.ok(existsSync(path.join(projectDir, "output", "report.md")));
});

test("the deck picks up the workspace brand, and a second workspace is independent", async (t) => {
  const dir = await tempDir(t);
  const acme = path.join(dir, "acme");
  const other = path.join(dir, "other");

  for (const root of [acme, other]) {
    await cli(["init", root]);
    await cli(["new", "deck", "--custom", "--workspace", root]);
  }
  await writeFile(path.join(acme, "design.yml"), "colors:\n  ink: AB1234\n", "utf8");

  await cli(["build", "--script", path.join(acme, "projects", "deck", "build.ts")]);
  await cli(["build", "--script", path.join(other, "projects", "deck", "build.ts")]);

  assert.ok(await deckMentions(path.join(acme, "projects", "deck", "output", "deck.pptx"), "AB1234"));
  assert.ok(
    !(await deckMentions(path.join(other, "projects", "deck", "output", "deck.pptx"), "AB1234")),
    "one workspace's brand must not leak into another"
  );
});

test("new rejects a deck id that is not kebab-case", async (t) => {
  const root = path.join(await tempDir(t), "decks");
  await cli(["init", root]);

  await assert.rejects(
    () => cli(["new", "Q4 Review", "--workspace", root]),
    (error: Error & { stderr?: string }) => {
      assert.match(error.stderr ?? "", /must be kebab-case/);
      return true;
    }
  );
});

/** True when any slide in the deck contains the given string. */
async function deckMentions(pptxPath: string, needle: string): Promise<boolean> {
  const { PptxPackage } = await import("./pptx-package.js");
  const { getSlideEntries } = await import("./ooxml.js");
  const pkg = await PptxPackage.load(pptxPath);
  for (const entry of await getSlideEntries(pkg)) {
    if ((await pkg.text(`ppt/slides/slide${entry.slideNumber}.xml`)).includes(needle)) return true;
  }
  return false;
}
