// Parsing and validation for `pptx-gen.config.yml`, the file that marks a
// folder as a pptx-gen workspace.
//
// A workspace holds the USER's data (brand, templates, decks). It is always a
// separate folder from the pptx-gen install, so the engine can be updated
// without touching anything the user owns.
import os from "node:os";
import path from "node:path";
import { readFileSync } from "node:fs";
import YAML from "yaml";

export const CONFIG_FILENAME = "pptx-gen.config.yml";

/** The config version this install understands. */
export const CONFIG_VERSION = 1;

export type WorkspaceConfig = {
  version: number;
  templates: string;
  projects: string;
  assets: string;
  design: string;
  designDoc: string;
};

/** Every field is optional in the file; these fill in the gaps. */
export const CONFIG_DEFAULTS: Omit<WorkspaceConfig, "version"> = {
  templates: "templates",
  projects: "projects",
  assets: "assets",
  design: "design.yml",
  designDoc: "design.md"
};

const PATH_KEYS = Object.keys(CONFIG_DEFAULTS) as (keyof typeof CONFIG_DEFAULTS)[];
const VALID_KEYS = ["version", ...PATH_KEYS];

export class WorkspaceConfigError extends Error {
  constructor(configPath: string, detail: string) {
    super(`${configPath}: ${detail}`);
    this.name = "WorkspaceConfigError";
  }
}

/** Expand a leading `~` so a config can point at `~/shared/templates`. */
export function expandHome(value: string): string {
  if (value === "~") return os.homedir();
  if (value.startsWith("~/")) return path.join(os.homedir(), value.slice(2));
  return value;
}

/**
 * Resolve a config value into an absolute path.
 *
 * Relative paths resolve against the directory holding the config file, NEVER
 * against process.cwd() — that cwd-dependence is exactly what this whole
 * feature exists to remove.
 */
export function resolveFromConfig(configPath: string, value: string): string {
  return path.resolve(path.dirname(configPath), expandHome(value));
}

export function parseWorkspaceConfig(raw: string, configPath: string): WorkspaceConfig {
  let parsed: unknown;
  try {
    parsed = YAML.parse(raw);
  } catch (error) {
    throw new WorkspaceConfigError(configPath, `could not parse YAML: ${(error as Error).message}`);
  }

  // An empty file is a valid, fully-default workspace.
  if (parsed === null || parsed === undefined) parsed = {};
  if (typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new WorkspaceConfigError(configPath, `expected a mapping of settings, got ${describe(parsed)}.`);
  }

  const record = parsed as Record<string, unknown>;

  for (const key of Object.keys(record)) {
    if (VALID_KEYS.includes(key)) continue;
    const suggestion = closestKey(key);
    throw new WorkspaceConfigError(
      configPath,
      `unknown key "${key}"${suggestion ? ` — did you mean "${suggestion}"?` : ""} Valid keys: ${VALID_KEYS.join(", ")}.`
    );
  }

  const version = record.version ?? CONFIG_VERSION;
  if (typeof version !== "number" || !Number.isInteger(version)) {
    throw new WorkspaceConfigError(configPath, `"version" must be an integer, got ${describe(version)}.`);
  }
  if (version !== CONFIG_VERSION) {
    throw new WorkspaceConfigError(
      configPath,
      `unsupported version ${version} (this install understands version ${CONFIG_VERSION}). Update pptx-gen.`
    );
  }

  const config = { version, ...CONFIG_DEFAULTS } as WorkspaceConfig;
  for (const key of PATH_KEYS) {
    const value = record[key];
    if (value === undefined || value === null) continue;
    if (typeof value !== "string" || value.trim() === "") {
      throw new WorkspaceConfigError(configPath, `"${key}" must be a non-empty string path, got ${describe(value)}.`);
    }
    config[key] = value;
  }

  return config;
}

export function readWorkspaceConfigSync(configPath: string): WorkspaceConfig {
  let raw: string;
  try {
    raw = readFileSync(configPath, "utf8");
  } catch (error) {
    throw new WorkspaceConfigError(configPath, `could not read the file: ${(error as Error).message}`);
  }
  return parseWorkspaceConfig(raw, configPath);
}

function describe(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "a list";
  return `a ${typeof value}`;
}

// A cheap did-you-mean, good enough for single-character slips like
// "template" -> "templates".
function closestKey(key: string): string | undefined {
  let best: string | undefined;
  let bestDistance = Infinity;
  for (const candidate of VALID_KEYS) {
    const distance = editDistance(key.toLowerCase(), candidate.toLowerCase());
    if (distance < bestDistance) {
      bestDistance = distance;
      best = candidate;
    }
  }
  return bestDistance <= 3 ? best : undefined;
}

function editDistance(a: string, b: string): number {
  const rows = a.length + 1;
  const cols = b.length + 1;
  let previous = Array.from({ length: cols }, (_unused, index) => index);
  for (let row = 1; row < rows; row += 1) {
    const current = [row, ...new Array<number>(cols - 1).fill(0)];
    for (let col = 1; col < cols; col += 1) {
      const cost = a[row - 1] === b[col - 1] ? 0 : 1;
      current[col] = Math.min(current[col - 1] + 1, previous[col] + 1, previous[col - 1] + cost);
    }
    previous = current;
  }
  return previous[cols - 1];
}
