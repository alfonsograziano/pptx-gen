import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { access } from "node:fs/promises";
import path from "node:path";

export const execFileAsync = promisify(execFile);

/**
 * Find the first usable executable from a list of candidates.
 *
 * A candidate that looks like a path (absolute, or containing a separator) is
 * checked directly on disk; a bare command name is resolved through `which` /
 * `where`.
 *
 * `envVar`, when set, is authoritative: it replaces the candidate list rather
 * than being prepended to it. Someone who names a binary explicitly wants that
 * binary, and quietly falling back to a different one on a typo would be worse
 * than reporting nothing found.
 */
export async function findExecutable(candidates: string[], envVar?: string): Promise<string | undefined> {
  const fromEnv = envVar ? process.env[envVar]?.trim() : undefined;
  for (const candidate of fromEnv ? [fromEnv] : candidates) {
    if (!candidate) continue;
    if (isPathLike(candidate)) {
      try {
        await access(candidate);
        return candidate;
      } catch {
        continue;
      }
    }
    try {
      const result = await execFileAsync(process.platform === "win32" ? "where" : "which", [candidate]);
      const commandPath = result.stdout.trim().split(/\r?\n/)[0];
      if (commandPath) return commandPath;
    } catch {}
  }
  return undefined;
}

// Windows absolute paths ("C:\Program Files\...") contain no forward slash, so
// testing for "/" alone would send them down the `where` branch and never find
// them.
function isPathLike(candidate: string): boolean {
  return path.isAbsolute(candidate) || candidate.includes("/") || candidate.includes("\\");
}

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
