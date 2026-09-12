---
name: customize-design
description: Customize the deck design system (colours, fonts, slide grid, logos) so generated decks match a specific brand. Use when the user wants to change the palette, set brand colours, swap fonts, add a logo, adjust the slide layout, or "make it look like our brand". Edits design.yml and design.md in the user's workspace, and optionally installs the new fonts. Trigger on requests like "use our brand colours", "change the accent to purple", "set the font to X", "add our logo", or "customise the design".
---

# Customize the design

## Purpose

Make every deck built by pptx-gen match a target brand by editing the design in
one place — the user's **workspace**, never the pptx-gen install.

## Resolve the workspace first

The brand lives in a workspace: a folder separate from the pptx-gen install, so
updating the engine never overwrites it. Get the real paths first:

```bash
pptx-gen workspace --json
```

Two files matter, and they must stay in sync:

- **`design`** (`design.yml`) — the values the engine reads: `colors`, `fonts`,
  `layout`, `logos`.
- **`designDoc`** (`design.md`) — the human-readable companion that documents
  the same values and the conventions. The other skills read this.

If the command fails with "No pptx-gen workspace found", stop and ask the user
where their workspace is, or offer `pptx-gen init <dir>`.

**Never edit `src/design.ts` in the pptx-gen install.** It holds the engine
defaults, it is shared by every workspace, and the user's next update will
overwrite whatever you put there.

## What the design controls (and what it does not)

The design system applies to **custom slides** and to **text/shapes the engine
adds** (added text, icons, custom layouts). It does **not** restyle the pixels of
an already-ingested cloned template: those slides keep the exact look they were
imported with. So changing the palette recolours custom slides and new elements,
not the baked-in appearance of the template library. If a brand needs
different-looking base templates, ingest templates that already match the brand.

## Steps

### 1. Gather the brand inputs

Ask for whatever is missing: primary dark colour, primary accent, one or two
secondary accents, the sans font, an optional serif for pull-quotes, an optional
mono for code, and any logo PNGs. If the user points at a brand guide, read it.

### 2. Edit `design.yml`

Every section and every key is optional — anything you leave out falls back to
the engine default, which is what lets the file survive an engine update. Change
values, not key names, so the rest of the code and the other skills keep working.

```yaml
colors:
  ink: 101820        # primary text and dark backgrounds
  accent: E8452C     # bars, highlights, key icons
  accent2: 7A3FF2    # diagrams only
fonts:
  sans: Inter        # the default for everything
  serif: Lora        # pull-quotes
  mono: JetBrains Mono
layout:
  LM: 0.75           # only if the brand really uses a different grid
```

- Colours are 6-digit hex. A leading `#` is accepted and stripped. Digits-only
  values like `000000` are fine unquoted.
- Valid colour keys: `ink`, `accent`, `white`, `accent2`, `accent3`, `surface`,
  `muted`, `faint`, `grey10`, `grey30`, `grey80`, `accentSoft`. An unknown key
  is an error naming the valid ones.
- Use exact Google Fonts family names where possible, so `install-fonts` can
  fetch them.
- Leave `logos` alone unless the user wants different file names; just drop the
  PNGs in (step 4).

### 3. Mirror the changes in `design.md`

Update the colour tables, the font roles, and any conventions so the doc matches
`design.yml`. The other skills read this file, so it must be accurate.

### 4. Add logos (optional)

Drop PNGs into the workspace's `assets` folder, named exactly as in the `logos`
section (`logo-mark-dark.png`, `logo-mark-light.png`, `logo-wordmark-dark.png`,
`logo-wordmark-light.png`). Use PNGs, not SVGs. Logos are workspace-only: if a
file is absent the helpers skip it silently, and nothing falls back to the
install, so a deck can never pick up a stray placeholder mark.

### 5. Install the fonts (optional)

Needed only for crisp local screenshots and native editing; generated decks embed
their own fonts. Run:

```bash
npm run install-fonts
```

If a family is not on Google Fonts, install it by hand and note it.

### 6. Verify

Check the workspace is healthy, then build a deck and look at it:

```bash
pptx-gen doctor
pptx-gen build --script <projects>/<some-deck>/build.ts
```

A malformed `design.yml` fails the build with a message naming the bad key — fix
it rather than working around it. Confirm custom slides pick up the new colours
and fonts, and report anything the user must do by hand (fonts not on Google
Fonts, missing logo files).
