// Creating and repairing a workspace.
//
// A workspace is the user's folder: brand, template library, decks. `init`
// seeds one; `doctor --fix` repairs one whose engine link or directories have
// gone missing (after a move, or a checkout without node_modules).
import path from "node:path";
import { existsSync } from "node:fs";
import { cp, mkdir, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { CONFIG_FILENAME } from "./workspace-config.js";
import { currentDesign } from "./design.js";
import { serializeDesign } from "./design-loader.js";
import { engineLinkPath, installDir, loadWorkspace, type Workspace, type WorkspaceProblem } from "./workspace.js";
import { ensureDir } from "./fs.js";

export type InitOptions = {
  dir: string;
  force?: boolean;
  starterTemplates?: boolean;
  link?: boolean;
};

export type InitResult = {
  workspace: Workspace;
  created: string[];
  skipped: string[];
};

const CONFIG_TEMPLATE = `# A pptx-gen workspace.
#
# This folder holds YOUR data — brand, templates, decks. The pptx-gen install
# is somewhere else entirely, so you can update the engine without touching
# anything here.
#
# Every path below is optional and relative to this file. Point one somewhere
# else to share it, for example:
#
#   templates: ~/shared/slide-templates
version: 1
templates: templates
projects: projects
assets: assets
design: design.yml
designDoc: design.md
customize: customize.md
`;

// Starts empty of rules on purpose: this is the user's file, and it fills up as
// they work. Every skill reads it before doing anything, and writes a new rule
// here whenever the user asks for something that will come up again.
const CUSTOMIZE_TEMPLATE = `# Customizations

House rules for this workspace. Every pptx-gen skill reads this file first and
follows what it says over its own defaults.

There are no rules yet. They arrive as you work: ask for something that will
come up again — "always do X after building", "never use Y" — and the agent
writes it here, so you only have to say it once.

## Rules

<!-- One bullet per rule. Say what to do, and when it applies. -->
`;

const GITIGNORE_TEMPLATE = `# The link to the pptx-gen engine, recreated by \`pptx-gen doctor --fix\`.
node_modules/

# Render artifacts. The .pptx in output/ is kept on purpose.
projects/*/output/screenshots/
projects/*/output/**/*.pdf
**/screenshots/*.pdf
templates-preview/
`;

/** Create (or repair, with `force`) a workspace. */
export async function initWorkspace(options: InitOptions): Promise<InitResult> {
  const root = path.resolve(options.dir);
  const created: string[] = [];
  const skipped: string[] = [];

  const configPath = path.join(root, CONFIG_FILENAME);
  if (existsSync(configPath) && !options.force) {
    throw new Error(
      `${configPath} already exists — this is already a workspace.\n` +
        "Run `pptx-gen init --force` to refresh its scaffolding, or pick another directory."
    );
  }

  await ensureDir(root);
  await writeIfAbsent(configPath, CONFIG_TEMPLATE, created, skipped, options.force);

  // The design is seeded from the engine's live defaults rather than a
  // committed copy, so it can never drift from them.
  await writeIfAbsent(path.join(root, "design.yml"), serializeDesign(currentDesign()), created, skipped, options.force);
  await writeIfAbsent(path.join(root, "customize.md"), CUSTOMIZE_TEMPLATE, created, skipped, options.force);
  await writeIfAbsent(path.join(root, ".gitignore"), GITIGNORE_TEMPLATE, created, skipped, options.force);
  await writeIfAbsent(
    path.join(root, "package.json"),
    `${JSON.stringify({ name: `${path.basename(root)}-decks`, private: true, type: "module" }, null, 2)}\n`,
    created,
    skipped,
    options.force
  );
  await writeIfAbsent(path.join(root, "tsconfig.json"), tsconfigTemplate(), created, skipped, options.force);

  const designDoc = path.join(installDir(), "starter", "design.md");
  if (existsSync(designDoc)) {
    await copyIfAbsent(designDoc, path.join(root, "design.md"), created, skipped, options.force);
  }

  for (const dir of ["templates", "projects", "assets"]) {
    const target = path.join(root, dir);
    if (existsSync(target)) skipped.push(target);
    else {
      await mkdir(target, { recursive: true });
      created.push(target);
    }
  }

  if (options.starterTemplates !== false) {
    await copyStarterTemplates(path.join(root, "templates"), created, skipped);
  }

  if (options.link !== false) {
    const link = await linkEngine(root);
    if (link) created.push(link);
  }

  return { workspace: loadWorkspace(root, "flag"), created, skipped };
}

/**
 * Point `<workspace>/node_modules/pptx-gen` at this install, so project scripts
 * can `import { Presentation } from "pptx-gen"`.
 *
 * The link names the install DIRECTORY, so updating the engine in place (a git
 * pull) needs no relinking. Only moving the install does.
 */
async function linkEngine(workspaceRoot: string): Promise<string | undefined> {
  const link = engineLinkPath(workspaceRoot);
  await ensureDir(path.dirname(link));
  // Replace rather than merge: a stale link pointing at an old checkout is
  // worse than no link, because it fails silently at the wrong version.
  await rm(link, { recursive: true, force: true });
  // "junction" is ignored on POSIX and avoids needing developer mode on Windows.
  await symlink(installDir(), link, "junction");
  return link;
}

/** Repair whatever `checkWorkspace` flagged as fixable. */
export async function fixWorkspace(workspace: Workspace, problems: WorkspaceProblem[]): Promise<string[]> {
  const fixed: string[] = [];

  for (const problem of problems) {
    if (problem.code === "missing-templates-dir") {
      await mkdir(workspace.templatesDir, { recursive: true });
      fixed.push(`Created ${workspace.templatesDir}`);
    } else if (problem.code === "missing-projects-dir") {
      await mkdir(workspace.projectsDir, { recursive: true });
      fixed.push(`Created ${workspace.projectsDir}`);
    } else if (problem.code === "missing-assets-dir") {
      await mkdir(workspace.assetsDir, { recursive: true });
      fixed.push(`Created ${workspace.assetsDir}`);
    } else if (problem.code === "missing-design") {
      await writeFile(workspace.designPath, serializeDesign(currentDesign()), "utf8");
      fixed.push(`Wrote ${workspace.designPath}`);
    } else if (problem.code === "missing-customize") {
      await writeFile(workspace.customizePath, CUSTOMIZE_TEMPLATE, "utf8");
      fixed.push(`Wrote ${workspace.customizePath}`);
    } else if (problem.code.endsWith("engine-link")) {
      const link = await linkEngine(workspace.root);
      fixed.push(`Linked ${link} -> ${installDir()}`);
    }
  }

  return fixed;
}

async function copyStarterTemplates(targetDir: string, created: string[], skipped: string[]): Promise<void> {
  const starterDir = path.join(installDir(), "starter", "templates");
  if (!existsSync(starterDir)) return;

  const { readdir } = await import("node:fs/promises");
  for (const entry of await readdir(starterDir, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const target = path.join(targetDir, entry.name);
    if (existsSync(target)) {
      skipped.push(target);
      continue;
    }
    await cp(path.join(starterDir, entry.name), target, { recursive: true });
    created.push(target);
  }
}

function tsconfigTemplate(): string {
  return `${JSON.stringify(
    {
      compilerOptions: {
        target: "ES2023",
        module: "NodeNext",
        moduleResolution: "NodeNext",
        strict: true,
        noEmit: true,
        allowImportingTsExtensions: true,
        skipLibCheck: true
      },
      include: ["projects/**/*.ts"]
    },
    null,
    2
  )}\n`;
}

async function writeIfAbsent(
  target: string,
  contents: string,
  created: string[],
  skipped: string[],
  force?: boolean
): Promise<void> {
  if (existsSync(target) && !force) {
    skipped.push(target);
    return;
  }
  // Never clobber a brand the user has edited, even under --force.
  if (existsSync(target) && force && (await readFile(target, "utf8")) !== contents) {
    skipped.push(target);
    return;
  }
  await writeFile(target, contents, "utf8");
  created.push(target);
}

async function copyIfAbsent(
  source: string,
  target: string,
  created: string[],
  skipped: string[],
  force?: boolean
): Promise<void> {
  if (existsSync(target)) {
    skipped.push(target);
    return;
  }
  void force;
  await cp(source, target);
  created.push(target);
}
