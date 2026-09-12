// Finding the files a slide refers to: icons, logos, and project images.
//
// Each kind has its own rule, and the differences are deliberate.
//
//   Icons   workspace first, then the set bundled with the engine. Copying
//           ~2000 Lucide files into every workspace would bloat it and freeze
//           the set at init time, so the workspace layer only shadows or adds.
//
//   Logos   workspace only, never a fallback. Falling back to the install
//           could stamp the tool's placeholder mark onto a client deck.
//
//   Project paths  relative to the deck's own folder, as before.
import path from "node:path";
import { existsSync } from "node:fs";

export type AssetResolver = {
  /**
   * Resolve an icon reference. A bare name like "rocket" is looked up in the
   * icon libraries; anything with a separator or an absolute path is treated
   * as a project file.
   */
  resolveIcon: (reference: string) => string;
  /** Resolve a logo file name; undefined when the workspace does not have it. */
  resolveLogo: (fileName: string) => string | undefined;
  /** Resolve a path written by the deck author, relative to the project. */
  resolveProjectPath: (filePath: string) => string;
  projectDir: string;
  assetsDir: string;
  iconDirs: string[];
};

export type AssetResolverOptions = {
  projectDir: string;
  /** The workspace's assets folder: logos, plus any icons the user added. */
  assetsDir: string;
  /** The engine's bundled icon set, searched after the workspace's. */
  bundledIconsDir?: string;
};

export class AssetNotFoundError extends Error {
  constructor(reference: string, searched: string[]) {
    super(`Icon "${reference}" was not found. Looked in:\n${searched.map((dir) => `  ${dir}`).join("\n")}`);
    this.name = "AssetNotFoundError";
  }
}

export function createAssetResolver(options: AssetResolverOptions): AssetResolver {
  const projectDir = path.resolve(options.projectDir);
  const assetsDir = path.resolve(options.assetsDir);
  const iconDirs = [path.join(assetsDir, "icons")];
  if (options.bundledIconsDir) {
    const bundled = path.resolve(options.bundledIconsDir);
    if (!iconDirs.includes(bundled)) iconDirs.push(bundled);
  }

  function resolveProjectPath(filePath: string): string {
    if (path.isAbsolute(filePath)) return filePath;
    return path.resolve(projectDir, filePath);
  }

  function resolveIcon(reference: string): string {
    // A path, not a name: the author is pointing at their own file.
    if (path.isAbsolute(reference) || reference.includes("/") || reference.includes("\\")) {
      return resolveProjectPath(reference);
    }

    const fileName = reference.endsWith(".svg") ? reference : `${reference}.svg`;
    for (const dir of iconDirs) {
      const candidate = path.join(dir, fileName);
      if (existsSync(candidate)) return candidate;
    }
    throw new AssetNotFoundError(reference, iconDirs);
  }

  function resolveLogo(fileName: string): string | undefined {
    const candidate = path.join(assetsDir, fileName);
    return existsSync(candidate) ? candidate : undefined;
  }

  return { resolveIcon, resolveLogo, resolveProjectPath, projectDir, assetsDir, iconDirs };
}
