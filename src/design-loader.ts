// Reading and writing a workspace `design.yml`.
//
// The file is a PARTIAL patch over the engine defaults in `design.ts`: any
// token it omits falls back to the default. That is what makes an engine
// update safe — a new version can add tokens without invalidating a brand file
// written months earlier.
import { readFileSync } from "node:fs";
import YAML from "yaml";
import {
  COLOR_NAMES,
  FONT_ROLES,
  LAYOUT_KEYS,
  LOGO_ROLES,
  type ColorName,
  type DesignPatch,
  type FontRole,
  type FullDesign,
  type LayoutSpec,
  type LogoRole
} from "./design-tokens.js";

const HEX = /^#?[0-9a-fA-F]{6}$/;

const SECTIONS = ["colors", "fonts", "layout", "logos"];

export class DesignFileError extends Error {
  constructor(designPath: string, detail: string) {
    super(`${designPath}: ${detail}`);
    this.name = "DesignFileError";
  }
}

export function parseDesignFile(raw: string, designPath: string): DesignPatch {
  // Parsed twice, on purpose.
  //
  // YAML's core schema turns a digits-only colour into a number, and does it
  // lossily: `000000` becomes 0, `001122` becomes 1122, `0x1122` becomes 4386.
  // Reconstructing the hex from those is guesswork, so colours, fonts, and
  // logo file names are read from a `failsafe` parse, where every scalar keeps
  // its literal source text. Layout still needs real numbers, so it comes from
  // the normal parse.
  const literal = parse(raw, designPath, { schema: "failsafe" });
  const typed = parse(raw, designPath, {});

  if (typed === undefined) return {};

  for (const key of Object.keys(typed)) {
    if (!SECTIONS.includes(key)) {
      throw new DesignFileError(designPath, `unknown section "${key}". Valid sections: ${SECTIONS.join(", ")}.`);
    }
  }

  const patch: DesignPatch = {};

  const colors = section(literal?.colors, "colors", designPath);
  if (colors) {
    const out: Partial<Record<ColorName, string>> = {};
    for (const [key, value] of Object.entries(colors)) {
      assertKnown(key, COLOR_NAMES, "colors", designPath);
      if (typeof value !== "string" || !HEX.test(value)) {
        throw new DesignFileError(
          designPath,
          `colors.${key} must be a 6-digit hex colour such as "3B82F6", got ${describe(value)}.`
        );
      }
      out[key as ColorName] = value.replace(/^#/, "").toUpperCase();
    }
    patch.colors = out;
  }

  const fonts = section(literal?.fonts, "fonts", designPath);
  if (fonts) {
    const out: Partial<Record<FontRole, string>> = {};
    for (const [key, value] of Object.entries(fonts)) {
      assertKnown(key, FONT_ROLES, "fonts", designPath);
      if (typeof value !== "string" || value.trim() === "") {
        throw new DesignFileError(designPath, `fonts.${key} must be a font family name, got ${describe(value)}.`);
      }
      out[key as FontRole] = value;
    }
    patch.fonts = out;
  }

  const layout = section(typed.layout, "layout", designPath);
  if (layout) {
    const out: Partial<LayoutSpec> = {};
    for (const [key, value] of Object.entries(layout)) {
      assertKnown(key, LAYOUT_KEYS, "layout", designPath);
      if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
        throw new DesignFileError(designPath, `layout.${key} must be a positive number of inches, got ${describe(value)}.`);
      }
      out[key as keyof LayoutSpec] = value;
    }
    patch.layout = out;
  }

  const logos = section(literal?.logos, "logos", designPath);
  if (logos) {
    const out: Partial<Record<LogoRole, string>> = {};
    for (const [key, value] of Object.entries(logos)) {
      assertKnown(key, LOGO_ROLES, "logos", designPath);
      if (typeof value !== "string" || value.trim() === "") {
        throw new DesignFileError(designPath, `logos.${key} must be a file name, got ${describe(value)}.`);
      }
      out[key as LogoRole] = value;
    }
    patch.logos = out;
  }

  return patch;
}

/**
 * Read a workspace design file.
 *
 * A missing file is not an error: it means "use the engine defaults". A
 * malformed one is, because silently ignoring a typo would ship an off-brand
 * deck.
 */
export function readDesignFileSync(designPath: string): DesignPatch | undefined {
  let raw: string;
  try {
    raw = readFileSync(designPath, "utf8");
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return undefined;
    throw new DesignFileError(designPath, `could not read the file: ${(error as Error).message}`);
  }
  return parseDesignFile(raw, designPath);
}

/**
 * Serialise a design to YAML, for seeding a new workspace's design.yml.
 *
 * Takes the design as an argument rather than reading it, so this module never
 * has to import `design.ts` — which would be a cycle, since `design.ts` imports
 * the reader from here to auto-load at startup.
 */
export function serializeDesign(design: FullDesign): string {
  const header = [
    "# Your brand.",
    "#",
    "# This file belongs to your workspace, not to the pptx-gen install, so",
    "# updating the engine never overwrites it.",
    "#",
    "# Every value is optional: delete any line to fall back to the engine",
    "# default. Colours are 6-digit hex without a leading '#'.",
    ""
  ].join("\n");
  return `${header}${YAML.stringify(design)}`;
}

function parse(raw: string, designPath: string, options: YAML.ParseOptions & YAML.SchemaOptions): Record<string, unknown> | undefined {
  let parsed: unknown;
  try {
    parsed = YAML.parse(raw, options);
  } catch (error) {
    throw new DesignFileError(designPath, `could not parse YAML: ${(error as Error).message}`);
  }
  if (parsed === null || parsed === undefined) return undefined;
  if (typeof parsed !== "object" || Array.isArray(parsed)) {
    throw new DesignFileError(designPath, `expected a mapping with any of ${SECTIONS.join(", ")}, got ${describe(parsed)}.`);
  }
  return parsed as Record<string, unknown>;
}

function section(value: unknown, name: string, designPath: string): Record<string, unknown> | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) {
    throw new DesignFileError(designPath, `"${name}" must be a mapping, got ${describe(value)}.`);
  }
  return value as Record<string, unknown>;
}

function assertKnown(key: string, valid: readonly string[], name: string, designPath: string): void {
  if (valid.includes(key)) return;
  throw new DesignFileError(designPath, `unknown key "${key}" in ${name}. Valid keys: ${valid.join(", ")}.`);
}

function describe(value: unknown): string {
  if (value === null) return "null";
  if (Array.isArray(value)) return "a list";
  if (typeof value === "string") return JSON.stringify(value);
  return `a ${typeof value}`;
}
