// The design system for generated decks.
//
// This is the single source of truth for colours, fonts, the slide grid, and
// optional logo assets. Custom slides, the deck engine, and the font installer
// all read from here. Edit this file (and `design.md`, its human-readable
// companion) to make every deck match your own brand.
//
// Colours are 6-digit hex WITHOUT a leading '#'. Keep the semantic key names
// (ink, accent, ...) and just change the values, so the rest of the code and
// the skills keep working.

export const C = {
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
} as const;

export type ColorName = keyof typeof C;

// Typography. `sans` is the default for everything; `serif` is reserved for
// pull-quotes and emphatic statements; `mono` is for code panels. These are the
// family names that must be embedded in templates and (optionally) installed
// locally for crisp screenshots. See `npm run install-fonts`.
export const FONTS = {
  sans: "Inter",
  serif: "Lora",
  mono: "JetBrains Mono"
} as const;

// The 16:9 slide grid, in inches. Every custom slide places content against
// these constants so decks stay aligned.
export const LAYOUT = {
  width: 10, // slide width (in)
  height: 5.625, // slide height (in): 16:9
  LM: 0.75, // left margin: all content starts here
  CW: 8.75, // content width: from LM to the right content edge
  BIL: 0.95, // bullet indent left
  LS: 1.3 // body line-spacing multiple
} as const;

// Optional logo assets, resolved relative to the `assets/` folder. Drop your
// own PNGs into `assets/` with these names to have them appear in slide footers
// and title slides. If a file is absent, the logo helpers skip it silently, so
// the tool works with no logos out of the box.
export const LOGO_FILES = {
  markDark: "logo-mark-dark.png", // small mark for LIGHT backgrounds
  markLight: "logo-mark-light.png", // small mark for DARK backgrounds
  wordmarkDark: "logo-wordmark-dark.png", // full wordmark for LIGHT backgrounds
  wordmarkLight: "logo-wordmark-light.png" // full wordmark for DARK backgrounds
} as const;
