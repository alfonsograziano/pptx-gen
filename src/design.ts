// The design system for generated decks.
//
// The values below are the engine's DEFAULTS. Your brand does not live here —
// it lives in `design.yml` in your workspace, which is loaded over these
// defaults when this module initialises. That split is deliberate: updating
// the engine must never overwrite your brand.
//
// Colours are 6-digit hex WITHOUT a leading '#'. Keep the semantic key names
// (ink, accent, ...) and just change the values in your `design.yml`, so the
// rest of the code and the skills keep working.
//
// These objects are MUTATED in place by `applyDesign` (see the note there), so
// that `import { C, LAYOUT } from "pptx-gen"` keeps working everywhere,
// including at module top level.
import { readDesignFileSync } from "./design-loader.js";
import { tryResolveWorkspaceSync } from "./workspace.js";
import type { ColorName, DesignPatch, FontRole, FullDesign, LayoutSpec, LogoRole } from "./design-tokens.js";

export type { ColorName, DesignPatch, FontRole, FullDesign, LayoutSpec, LogoRole } from "./design-tokens.js";
export { COLOR_NAMES, FONT_ROLES, LAYOUT_KEYS, LOGO_ROLES } from "./design-tokens.js";

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
  accentSoft: "E8F0FE",

  // Raised surfaces. A card that should sit *above* the page needs a tone
  // separate from the page itself, and the two directions are not symmetric:
  // on light the card lifts towards white, on dark it lifts away from black.
  // In this default palette the page is already pure white, so `paperSoft`
  // matches `white`; a brand whose paper is off-white gives them different
  // values.
  paperSoft: "FFFFFF", // raised card on a LIGHT slide
  inkSoft: "1C2338", // raised panel on a DARK slide

  // The accent, retuned for dark backgrounds. One accent cannot clear the
  // contrast bar on both paper and ink, so brands that use dark slides set
  // this to a lighter cut of the same hue.
  accentOnDark: "60A5FA"
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

/** The current live design, with every token present. Seeds a new `design.yml`. */
export function currentDesign(): FullDesign {
  return {
    colors: { ...C },
    fonts: { ...FONTS },
    layout: { ...LAYOUT },
    logos: { ...LOGO_FILES }
  };
}

/** Where the live design came from. Reported by `pptx-gen workspace`. */
export type DesignOrigin = { file: string } | { defaults: true };

let origin: DesignOrigin = { defaults: true };

export function designOrigin(): DesignOrigin {
  return origin;
}

/**
 * Load the workspace design over the defaults, as a side effect of importing
 * this module.
 *
 * Doing it here, synchronously, is what makes top-level reads safe: ESM
 * guarantees this module finishes initialising before any importer's body
 * runs, so a project's `const { LM } = LAYOUT` already sees workspace values.
 *
 * Best-effort by design. No workspace, or a workspace with no design.yml, just
 * means the engine defaults — the engine has to work in a bare checkout. A
 * malformed design.yml does throw, because silently ignoring it would ship an
 * off-brand deck.
 *
 * Set PPTX_GEN_NO_AUTOLOAD=1 to pin the defaults (the test suite does).
 */
function autoLoadDesign(): void {
  if (process.env.PPTX_GEN_NO_AUTOLOAD === "1") return;

  const workspace = tryResolveWorkspaceSync();
  if (!workspace) return;

  const patch = readDesignFileSync(workspace.designPath);
  if (!patch) return;

  applyDesign(patch);
  origin = { file: workspace.designPath };
}

autoLoadDesign();
