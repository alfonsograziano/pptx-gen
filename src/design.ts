// The design system for generated decks.
//
// The values below are the engine's DEFAULTS. Your brand does not live here —
// it lives in `design.yml` in your workspace, which is loaded over these
// defaults at startup. That split is deliberate: updating the engine must
// never overwrite your brand.
//
// Colours are 6-digit hex WITHOUT a leading '#'. Keep the semantic key names
// (ink, accent, ...) and just change the values in your `design.yml`, so the
// rest of the code and the skills keep working.
//
// These objects are MUTATED in place by `applyDesign` (see the note there), so
// that `import { C, LAYOUT } from "pptx-gen"` keeps working everywhere,
// including at module top level.

export type ColorName =
  | "ink"
  | "accent"
  | "white"
  | "accent2"
  | "accent3"
  | "surface"
  | "muted"
  | "faint"
  | "grey10"
  | "grey30"
  | "grey80"
  | "accentSoft";

export type FontRole = "sans" | "serif" | "mono";

export type LogoRole = "markDark" | "markLight" | "wordmarkDark" | "wordmarkLight";

export type LayoutSpec = {
  /** Slide width (in). */
  width: number;
  /** Slide height (in): 16:9. */
  height: number;
  /** Left margin: all content starts here. */
  LM: number;
  /** Content width: from LM to the right content edge. */
  CW: number;
  /** Bullet indent left. */
  BIL: number;
  /** Body line-spacing multiple. */
  LS: number;
};

export const COLOR_NAMES: readonly ColorName[] = [
  "ink", "accent", "white", "accent2", "accent3",
  "surface", "muted", "faint", "grey10", "grey30", "grey80", "accentSoft"
];

export const FONT_ROLES: readonly FontRole[] = ["sans", "serif", "mono"];

export const LOGO_ROLES: readonly LogoRole[] = ["markDark", "markLight", "wordmarkDark", "wordmarkLight"];

export const LAYOUT_KEYS: readonly (keyof LayoutSpec)[] = ["width", "height", "LM", "CW", "BIL", "LS"];

export const C: Record<ColorName, string> = {
  // Core pairing: a dark tone for text and dark backgrounds, plus one vivid
  // accent for bars, highlights, and calls to action.
  ink: "12182B", // primary text, dark hero/section backgrounds
  accent: "3B82F6", // primary accent: bars, highlights, key icons
  white: "FFFFFF", // page/slide backgrounds, text on dark backgrounds

  // Secondary accents for diagrams and infographics only.
  accent2: "8B5CF6", // violet
  accent3: "0EA5E9", // sky

  // Neutral tints for low-contrast backgrounds and supporting text.
  surface: "F5F7FA", // faint panel/background fill
  muted: "5B6472", // muted supporting text (captions, footers)
  faint: "9AA3B2", // faint lines and de-emphasised labels

  // Greys for borders and separators.
  grey10: "EEF0F3",
  grey30: "D5D9E0",
  grey80: "3A3F4B",

  // Soft accent tint (e.g. a highlighted card fill).
  accentSoft: "E8F0FE"
};

// Typography. `sans` is the default for everything; `serif` is reserved for
// pull-quotes and emphatic statements; `mono` is for code panels. These are the
// family names that must be embedded in templates and (optionally) installed
// locally for crisp screenshots. See `pptx-gen fonts install`.
export const FONTS: Record<FontRole, string> = {
  sans: "Inter",
  serif: "Lora",
  mono: "JetBrains Mono"
};

// The 16:9 slide grid, in inches. Every custom slide places content against
// these constants so decks stay aligned.
export const LAYOUT: LayoutSpec = {
  width: 10,
  height: 5.625,
  LM: 0.75,
  CW: 8.75,
  BIL: 0.95,
  LS: 1.3
};

// Optional logo assets, resolved relative to the workspace's `assets/` folder.
// Drop your own PNGs there with these names to have them appear in slide
// footers and title slides. If a file is absent, the logo helpers skip it
// silently, so the tool works with no logos out of the box.
export const LOGO_FILES: Record<LogoRole, string> = {
  markDark: "logo-mark-dark.png", // small mark for LIGHT backgrounds
  markLight: "logo-mark-light.png", // small mark for DARK backgrounds
  wordmarkDark: "logo-wordmark-dark.png", // full wordmark for LIGHT backgrounds
  wordmarkLight: "logo-wordmark-light.png" // full wordmark for DARK backgrounds
};

/** A partial override of the design, as read from a workspace `design.yml`. */
export type DesignPatch = {
  colors?: Partial<Record<ColorName, string>>;
  fonts?: Partial<Record<FontRole, string>>;
  layout?: Partial<LayoutSpec>;
  logos?: Partial<Record<LogoRole, string>>;
};

/**
 * Apply a workspace design over the engine defaults, in place.
 *
 * The objects are mutated rather than replaced so that every existing
 * `import { C } from "pptx-gen"` — including a top-level
 * `const { LM, CW } = LAYOUT` in a project's `custom.ts` — sees the workspace
 * values without any call-site changes.
 *
 * ORDERING RULE: this must run before the first `await import()` of any project
 * script. Anything that destructures these objects captures the values it sees
 * at that moment, so a late re-apply would leave stale copies behind. Never
 * call it per-deck or inside `render()`.
 */
export function applyDesign(patch: DesignPatch): void {
  if (patch.colors) Object.assign(C, patch.colors);
  if (patch.fonts) Object.assign(FONTS, patch.fonts);
  if (patch.layout) Object.assign(LAYOUT, patch.layout);
  if (patch.logos) Object.assign(LOGO_FILES, patch.logos);
}

/** The current live design, as a patch. Used to seed a new `design.yml`. */
export function currentDesign(): Required<DesignPatch> {
  return {
    colors: { ...C },
    fonts: { ...FONTS },
    layout: { ...LAYOUT },
    logos: { ...LOGO_FILES }
  };
}
