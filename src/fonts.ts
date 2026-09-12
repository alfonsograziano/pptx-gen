import { readdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

// Font locations shared by the font installer and the figure renderer.
//
// This lives apart from `install-fonts.ts` on purpose: that module runs its
// `main()` on import, so importing it to reuse a path helper would download and
// install fonts as a side effect.

/** Where the OS looks for user-installed fonts. */
export function userFontDir(): string {
  if (process.platform === "darwin") return path.join(os.homedir(), "Library", "Fonts");
  if (process.platform === "win32") {
    const localAppData = process.env.LOCALAPPDATA ?? path.join(os.homedir(), "AppData", "Local");
    return path.join(localAppData, "Microsoft", "Windows", "Fonts");
  }
  return path.join(os.homedir(), ".local", "share", "fonts");
}

function systemFontDirs(): string[] {
  if (process.platform === "darwin") return ["/Library/Fonts", "/System/Library/Fonts"];
  if (process.platform === "win32") return [path.join(process.env.WINDIR ?? "C:\\Windows", "Fonts")];
  return ["/usr/share/fonts", "/usr/local/share/fonts"];
}

/**
 * Is a font family installed on this machine?
 *
 * Deliberately a filename check rather than a real font query: `fc-list` and
 * `system_profiler SPFontsDataType` take one to three seconds, which would cost
 * more than the render this is guarding. The answer only drives a warning, so
 * being approximate is the right trade.
 */
export async function isFontInstalled(family: string): Promise<boolean> {
  const needle = family.toLowerCase().replace(/[^a-z0-9]/g, "");
  if (!needle) return true;
  for (const dir of [userFontDir(), ...systemFontDirs()]) {
    const entries = await readdir(dir).catch(() => []);
    if (
      entries.some((entry) =>
        entry
          .toLowerCase()
          .replace(/[^a-z0-9]/g, "")
          .includes(needle)
      )
    ) {
      return true;
    }
  }
  return false;
}
