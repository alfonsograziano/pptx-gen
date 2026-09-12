// The vocabulary of the design system: which tokens exist, and their types.
//
// This module deliberately imports nothing. `design.ts` (the live values) and
// `design-loader.ts` (reading a workspace design.yml) both depend on it, which
// keeps those two from depending on each other.

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
  | "accentSoft"
  | "paperSoft"
  | "inkSoft"
  | "accentOnDark";

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

/** A partial override of the design, as read from a workspace `design.yml`. */
export type DesignPatch = {
  colors?: Partial<Record<ColorName, string>>;
  fonts?: Partial<Record<FontRole, string>>;
  layout?: Partial<LayoutSpec>;
  logos?: Partial<Record<LogoRole, string>>;
};

/** A design with every token present. */
export type FullDesign = Required<DesignPatch>;

export const COLOR_NAMES: readonly ColorName[] = [
  "ink", "accent", "white", "accent2", "accent3",
  "surface", "muted", "faint", "grey10", "grey30", "grey80", "accentSoft",
  "paperSoft", "inkSoft", "accentOnDark"
];

export const FONT_ROLES: readonly FontRole[] = ["sans", "serif", "mono"];

export const LOGO_ROLES: readonly LogoRole[] = ["markDark", "markLight", "wordmarkDark", "wordmarkLight"];

export const LAYOUT_KEYS: readonly (keyof LayoutSpec)[] = ["width", "height", "LM", "CW", "BIL", "LS"];
