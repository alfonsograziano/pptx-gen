// Where the user's data lives, and how we find it.
//
// The pptx-gen INSTALL (this repo) holds the engine and a few seed files. The
// user's brand, template library, and decks live in a WORKSPACE: any folder
// containing a `pptx-gen.config.yml`. The two are always separate, so the
// engine can be updated without carrying decks along, and one install can
// serve several workspaces (one per client or brand).
import os from "node:os";
import path from "node:path";
import { existsSync, realpathSync, statSync } from "node:fs";
import { fileURLToPath } from "node:url";
import {
  CONFIG_FILENAME,
  readWorkspaceConfigSync,
  resolveFromConfig,
  type WorkspaceConfig
} from "./workspace-config.js";

export type WorkspaceSource = "flag" | "env" | "discovered" | "home";

export type Workspace = {
  /** Absolute path to the workspace root (the folder holding the config). */
  root: string;
  configPath: string;
  config: WorkspaceConfig;
  templatesDir: string;
  projectsDir: string;
  assetsDir: string;
  /** `<assetsDir>/icons` — user icons, which shadow the bundled set. */
  iconsDir: string;
  designPath: string;
  designDocPath: string;
  /** `customize.md` — this workspace's house rules, read by every skill. */
  customizePath: string;
  /** Where the engine is installed. Never written to. */
  installDir: string;
  /** The Lucide set that ships with the engine, used as an icon fallback. */
  bundledIconsDir: string;
  source: WorkspaceSource;
};

export type ResolveOptions = {
  /** An explicit workspace directory (the `--workspace` flag). Wins over all else. */
  explicit?: string;
  cwd?: string;
  /** Directory of the entry script; defaults to the dirname of process.argv[1]. */
  entryDir?: string;
  env?: NodeJS.ProcessEnv;
};

export const WORKSPACE_ENV_VAR = "PPTX_GEN_WORKSPACE";

export class WorkspaceNotFoundError extends Error {
  readonly searched: string[];

  constructor(searched: string[]) {
    super(
      [
        "No pptx-gen workspace found.",
        "",
        `Looked for ${CONFIG_FILENAME} in:`,
        ...searched.map((dir) => `  ${dir}`),
        "",
        "A workspace holds your brand, templates, and decks — separate from the",
        "pptx-gen install, so you can update the engine without touching them.",
        "",
        "Create one:      pptx-gen init ~/decks",
        "Or point at one: pptx-gen --workspace ~/decks <command>",
        `                 export ${WORKSPACE_ENV_VAR}=~/decks`
      ].join("\n")
    );
    this.name = "WorkspaceNotFoundError";
    this.searched = searched;
  }
}

const HERE = path.dirname(fileURLToPath(import.meta.url));

/** The root of the pptx-gen install (the folder holding `src/` and `package.json`). */
export function installDir(): string {
  return path.resolve(HERE, "..");
}

/** `~/.pptx-gen`, the fallback workspace when nothing else is configured. */
export function homeWorkspaceDir(): string {
  return path.join(os.homedir(), ".pptx-gen");
}

/** The symlink that lets a project script `import { ... } from "pptx-gen"`. */
export function engineLinkPath(workspaceRoot: string): string {
  return path.join(workspaceRoot, "node_modules", "pptx-gen");
}

/**
 * Resolve the workspace, or throw a WorkspaceNotFoundError listing where we
 * looked.
 *
 * Precedence: explicit flag, env var, nearest config walking up (from the entry
 * script first, then the cwd), then `~/.pptx-gen`.
 */
export function resolveWorkspaceSync(options: ResolveOptions = {}): Workspace {
  const found = findWorkspaceRoot(options);
  if ("workspace" in found) return found.workspace;
  throw new WorkspaceNotFoundError(found.searched);
}

/**
 * Like resolveWorkspaceSync, but returns undefined instead of throwing.
 *
 * Used where a missing workspace must not be fatal — notably the design
 * auto-load, which falls back to the engine's built-in defaults.
 */
export function tryResolveWorkspaceSync(options: ResolveOptions = {}): Workspace | undefined {
  try {
    const found = findWorkspaceRoot(options);
    return "workspace" in found ? found.workspace : undefined;
  } catch {
    // A malformed config must not break design loading; the CLI surfaces the
    // real error when it resolves the workspace for a command that needs it.
    return undefined;
  }
}

/** Build a Workspace from a directory known to contain a config file. */
export function loadWorkspace(root: string, source: WorkspaceSource): Workspace {
  const absoluteRoot = path.resolve(root);
  const configPath = path.join(absoluteRoot, CONFIG_FILENAME);
  const config = readWorkspaceConfigSync(configPath);
  const assetsDir = resolveFromConfig(configPath, config.assets);
  const install = installDir();

  return {
    root: absoluteRoot,
    configPath,
    config,
    templatesDir: resolveFromConfig(configPath, config.templates),
    projectsDir: resolveFromConfig(configPath, config.projects),
    assetsDir,
    iconsDir: path.join(assetsDir, "icons"),
    designPath: resolveFromConfig(configPath, config.design),
    designDocPath: resolveFromConfig(configPath, config.designDoc),
    customizePath: resolveFromConfig(configPath, config.customize),
    installDir: install,
    bundledIconsDir: path.join(install, "assets", "icons"),
    source
  };
}

type FindResult = { workspace: Workspace } | { searched: string[] };

function findWorkspaceRoot(options: ResolveOptions): FindResult {
  const env = options.env ?? process.env;
  const searched: string[] = [];

  const explicit = options.explicit?.trim();
  if (explicit) {
    const root = path.resolve(expandTilde(explicit));
    if (hasConfig(root)) return { workspace: loadWorkspace(root, "flag") };
    searched.push(root);
    // An explicit pointer that misses is always an error — never silently fall
    // through to some other workspace and write a deck into the wrong brand.
    return { searched };
  }

  const fromEnv = env[WORKSPACE_ENV_VAR]?.trim();
  if (fromEnv) {
    const root = path.resolve(expandTilde(fromEnv));
    if (hasConfig(root)) return { workspace: loadWorkspace(root, "env") };
    searched.push(root);
    return { searched };
  }

  // Walk up from the entry script first, so `tsx /ws/projects/deck/build.ts`
  // finds /ws no matter which directory it was launched from.
  for (const start of startDirs(options)) {
    const hit = walkUp(start, searched);
    if (hit) return { workspace: loadWorkspace(hit, "discovered") };
  }

  const home = homeWorkspaceDir();
  if (hasConfig(home)) return { workspace: loadWorkspace(home, "home") };
  searched.push(home);

  return { searched };
}

function startDirs(options: ResolveOptions): string[] {
  const dirs: string[] = [];

  const entryDir = options.entryDir ?? (process.argv[1] ? path.dirname(path.resolve(process.argv[1])) : undefined);
  if (entryDir) dirs.push(entryDir);

  const cwd = path.resolve(options.cwd ?? process.cwd());
  if (!dirs.includes(cwd)) dirs.push(cwd);

  return dirs;
}

function walkUp(start: string, searched: string[]): string | undefined {
  let dir = start;
  for (;;) {
    if (!searched.includes(dir)) searched.push(dir);
    if (hasConfig(dir)) return dir;
    const parent = path.dirname(dir);
    if (parent === dir) return undefined;
    dir = parent;
  }
}

function hasConfig(dir: string): boolean {
  try {
    return statSync(path.join(dir, CONFIG_FILENAME)).isFile();
  } catch {
    return false;
  }
}

function expandTilde(value: string): string {
  if (value === "~") return os.homedir();
  if (value.startsWith("~/")) return path.join(os.homedir(), value.slice(2));
  return value;
}

export type WorkspaceProblem = {
  code: string;
  message: string;
  /** True when `pptx-gen doctor --fix` can repair it. */
  fixable: boolean;
};

/** The checks behind `pptx-gen doctor`. */
export async function checkWorkspace(workspace: Workspace): Promise<WorkspaceProblem[]> {
  const problems: WorkspaceProblem[] = [];

  const dirs: [string, string][] = [
    ["templates", workspace.templatesDir],
    ["projects", workspace.projectsDir],
    ["assets", workspace.assetsDir]
  ];
  for (const [name, dir] of dirs) {
    if (!existsSync(dir)) {
      problems.push({ code: `missing-${name}-dir`, message: `Missing ${name} directory: ${dir}`, fixable: true });
    }
  }

  if (!existsSync(workspace.designPath)) {
    problems.push({
      code: "missing-design",
      message: `Missing design file: ${workspace.designPath} (the engine defaults are being used)`,
      fixable: true
    });
  }

  if (!existsSync(workspace.customizePath)) {
    problems.push({
      code: "missing-customize",
      message: `Missing customizations file: ${workspace.customizePath} (this workspace has no house rules)`,
      fixable: true
    });
  }

  problems.push(...checkEngineLink(workspace));

  return problems;
}

export function checkEngineLink(workspace: Workspace): WorkspaceProblem[] {
  const link = engineLinkPath(workspace.root);
  if (!existsSync(link)) {
    return [{
      code: "missing-engine-link",
      message: `Missing engine link: ${link}. Project scripts that import "pptx-gen" will fail.`,
      fixable: true
    }];
  }

  try {
    const target = realpathSync(link);
    const expected = realpathSync(workspace.installDir);
    if (target !== expected) {
      return [{
        code: "stale-engine-link",
        message: `Engine link points at ${target}, but this install is ${expected}.`,
        fixable: true
      }];
    }
  } catch (error) {
    return [{
      code: "broken-engine-link",
      message: `Engine link is broken: ${(error as Error).message}`,
      fixable: true
    }];
  }

  return [];
}
