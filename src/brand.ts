export const C = {
  midnight: "000E38",
  green: "00E5A4",
  blue: "166BFF",
  purple: "8950FF",
  white: "FFFFFF",
  grey10: "EAEBED",
  grey30: "D9D9D9",
  grey80: "454551",
  mid10: "F4F8FA",
  mid30: "97A1B8",
  mid50: "526288",
  green10: "E5FCF5",
  purple10: "F3EDFF",
  blue10: "E8F0FF"
} as const;

export const LAYOUT = {
  width: 10,
  height: 5.625,
  LM: 0.75,
  CW: 8.75,
  BIL: 0.95,
  LS: 1.3
} as const;

export const LOGO_FILES = {
  nMark: "NF_Deck_Symbol_Reverse.png",
  nMarkColor: "NF_Deck_Symbol_Color.png",
  wordmarkDark: "NF_Deck_Logo_Reversed.png",
  wordmarkLight: "NF_Deck_Logo_Color.png"
} as const;

export type NearformColorName = keyof typeof C;
