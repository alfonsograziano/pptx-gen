---
name: customize-design
description: Customize the deck design system (colours, fonts, slide grid, logos) so generated decks match a specific brand. Use when the user wants to change the palette, set brand colours, swap fonts, add a logo, adjust the slide layout, or "make it look like our brand". Edits design.md and src/design.ts together, and optionally installs the new fonts. Trigger on requests like "use our brand colours", "change the accent to purple", "set the font to X", "add our logo", or "customise the design".
---

# Customize the design

## Purpose

Make every deck built by pptx-gen match a target brand by editing the design
system in one place. There are two files, and they must stay in sync:

- `src/design.ts` — the code source of truth: `C` (colours), `FONTS`, `LAYOUT`,
  `LOGO_FILES`. This is what the engine reads.
- `design.md` — the human-readable companion that documents the same values and
  the conventions.

## What the design controls (and what it does not)

The design system applies to **custom slides** and to **text/shapes the engine
adds** (added text, icons, custom layouts). It does **not** restyle the pixels of
an already-ingested cloned template: those slides keep the exact look they were
imported with. So changing the palette recolours custom slides and new elements,
not the baked-in appearance of `templates/*`. If a brand needs different-looking
base templates, ingest templates that already match the brand.

## Steps

### 1. Gather the brand inputs

Ask for whatever is missing: primary dark colour, primary accent, one or two
secondary accents, the sans font, an optional serif for pull-quotes, an optional
mono for code, and any logo PNGs. If the user points at a brand guide, read it.

### 2. Edit `src/design.ts`

- Update `C` values (6-digit hex, no leading `#`). Keep the semantic key names
  (`ink`, `accent`, `accent2`, ...) so the rest of the code and the skills keep
  working. Change values, not keys.
- Update `FONTS.sans`, `FONTS.serif`, `FONTS.mono` to the brand families. Use
  exact Google Fonts family names where possible, so `install-fonts` can fetch
  them.
- Adjust `LAYOUT` only if the brand uses a different grid (most do not).
- Leave `LOGO_FILES` names as they are; just drop the PNGs into `assets/` (see
  step 4).

### 3. Mirror the changes in `design.md`

Update the colour tables, the font roles, and any conventions so the doc matches
`src/design.ts`. The skills read `design.md`, so it must be accurate.

### 4. Add logos (optional)

Drop PNGs into `assets/` named exactly as in `LOGO_FILES`
(`logo-mark-dark.png`, `logo-mark-light.png`, `logo-wordmark-dark.png`,
`logo-wordmark-light.png`). Use PNGs, not SVGs. If a file is absent, the logo
helpers skip it silently.

### 5. Install the fonts (optional)

Needed only for crisp local screenshots and native editing; generated decks embed
their own fonts. Run:

```bash
npm run install-fonts
```

If a family is not on Google Fonts, install it by hand and note it.

### 6. Verify

Typecheck, then build a deck (or the example) and inspect it:

```bash
npm run build
npm run self-validate
```

Confirm custom slides pick up the new colours and fonts, and report anything the
user must do by hand (fonts not on Google Fonts, missing logo files).
