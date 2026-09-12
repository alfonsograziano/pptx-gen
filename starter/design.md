# Design system

This is the design system for decks built with pptx-gen — a clean, neutral
starting point you are meant to change.

The values live in `design.yml`, next to this file in your workspace. That file
is yours: it belongs to the workspace, not to the pptx-gen install, so updating
the engine never overwrites your brand. Anything you leave out of it falls back
to the engine's defaults, which is what keeps it working across updates.

Edit `design.yml` and this document together (or run the `customize-design`
skill) to make every deck match your brand. Do not edit `src/design.ts` in the
pptx-gen install: that is shared by every workspace and is overwritten by
updates.

> **One rule worth keeping: sentence case.** Titles, headings, labels, and
> buttons read best in sentence case: only the first word and proper nouns are
> capitalised ("How we ship faster", not "How We Ship Faster"). It reads as more
> human and is easier to translate. This is a convention, not a hard constraint.

---

## Colours

Colours live under `colors:` in `design.yml`, as 6-digit hex without a leading `#`.

### Core

| Token | Default | Usage |
|---|---|---|
| `ink` | `#12182B` | Primary text, dark hero/section backgrounds |
| `accent` | `#3B82F6` | Accent bars, highlights, CTAs, key icons |
| `white` | `#FFFFFF` | Page/slide backgrounds, text on dark backgrounds |

### Secondary accents (diagrams and infographics only)

| Token | Default |
|---|---|
| `accent2` | `#8B5CF6` |
| `accent3` | `#0EA5E9` |

### Neutrals

| Token | Default | Usage |
|---|---|---|
| `surface` | `#F5F7FA` | Faint panel/background fill |
| `muted` | `#5B6472` | Muted supporting text (captions, footers) |
| `faint` | `#9AA3B2` | Faint lines, de-emphasised labels |
| `grey10` | `#EEF0F3` | Light separators |
| `grey30` | `#D5D9E0` | Borders, table rules |
| `grey80` | `#3A3F4B` | Strong borders |
| `accentSoft` | `#E8F0FE` | Soft accent tint (highlighted card fill) |

### Colour principles

- **`ink` + `accent` are the core pairing.** Use them for any hero, cover, or key
  content area.
- **White is the default background.** Pair it with `ink` text.
- **Use `accent2` and `accent3` only in diagrams and infographics**, to add
  variety without replacing the core pairing.
- **Mind contrast.** A vivid accent behind small text can fail accessibility.
  Keep readable body text as `ink` on white or white on `ink`. Use the accent for
  bars, icons, and decorative fills, not for long runs of small text.

---

## Typography

Font families live under `fonts:` in `design.yml`.

| Role | Default | When to use |
|---|---|---|
| `sans` | Inter | Everything: titles, headings, body, labels, UI text |
| `serif` | Lora | Pull-quotes and emphatic statements, used sparingly |
| `mono` | JetBrains Mono | Code panels |

### Sizes

| Level | Size (pt) | Notes |
|---|---|---|
| Headline | 14 or 16 | Pick one and use it consistently across the whole deck |
| Subheading | 12 | Section headers, slide subheadings |
| Body | 10 | Paragraphs, bullets, labels |
| Pull-quote | 14 or 16 | `serif`, for callouts and quotes only |

### Principles

- **`sans` is the default.** When in doubt, use it. Reserve `serif` for warmth or
  emphasis; never use it for standard body paragraphs or bullets.
- Do not bold headings; the font weight carries hierarchy.
- Keep body line height at `LAYOUT.LS` (1.3).
- Use at most two families per deck (`sans` + `serif`). `mono` is for code only.

---

## Slide grid

Layout constants live under `layout:` in `design.yml`, in inches.

| Property | Value |
|---|---|
| Slide size | 10 × 5.625 in (16:9) |
| Left margin (`LM`) | 0.75 in — all content starts here |
| Content width (`CW`) | 8.75 in — from `LM` to the right content edge |
| Bullet indent left (`BIL`) | 0.95 in |
| Body line spacing (`LS`) | 1.3 |

---

## Slide conventions

These are conventions the custom-slide helpers follow. They are easy to change,
but they give a deck a consistent rhythm.

- **Backgrounds.** White for standard content slides; `ink` for covers, section
  breaks, and closing slides.
- **Header.** Content slides open with a header top-left: `sans` 14pt, `ink` on
  light backgrounds and white on dark. Covers and closing slides skip it.
- **Footer.** A page number bottom-left in `muted` (`sans` 7pt) plus an optional
  logo mark bottom-right. The number is a live PowerPoint slide-number field, not
  text: nobody passes it in, and it stays right when slides are reordered. A deck
  that opens on an unnumbered cover counts from zero, so the first slide showing a
  footer reads "1".

---

## Logos (optional)

Logo file names live under `logos:` in `design.yml`. Drop PNGs with these names
into your workspace's `assets/` folder to have them appear automatically:

| File | Usage |
|---|---|
| `logo-mark-dark.png` | Small mark for light backgrounds |
| `logo-mark-light.png` | Small mark for dark backgrounds |
| `logo-wordmark-dark.png` | Full wordmark for light backgrounds |
| `logo-wordmark-light.png` | Full wordmark for dark backgrounds |

If a file is missing, the logo helpers skip it silently. The tool works with no
logos out of the box. Use PNGs, not SVGs: SVG logos do not embed reliably in
PowerPoint.

---

## Figures

A [figure](figure-instructions.md) is an HTML file rendered to a picture at build
time, for content that has no native form: a UI mockup, a chart, a rendered
document. It must look like it belongs on the slide, so the engine injects every
token below as a CSS custom property on `:root`:

| Token in `src/design.ts` | CSS variable |
|---|---|
| `C.ink` | `--ink` |
| `C.accent` | `--accent` |
| `C.accent2` / `C.accent3` | `--accent-2` / `--accent-3` |
| `C.white` | `--white` |
| `C.surface` | `--surface` |
| `C.muted` | `--muted` |
| `C.faint` | `--faint` |
| `C.grey10` / `C.grey30` / `C.grey80` | `--grey-10` / `--grey-30` / `--grey-80` |
| `C.accentSoft` | `--accent-soft` |
| `FONTS.sans` / `serif` / `mono` | `--font-sans` / `--font-serif` / `--font-mono` |

The list is generated from `C` and `FONTS`, so a token added there is available
in figures with no further work. Figure markup uses **only** these variables: a
literal hex in a figure is an off-brand colour that will not follow a rebrand.

Figures need the design fonts installed locally (`npm run install-fonts`). Unlike
a slide, a figure cannot embed a font — it bakes in whatever the build machine
had, so the same deck built elsewhere would look different.

---

## Quick checklist

Before shipping a deck, verify:

- [ ] Text is sentence case
- [ ] `sans` for headings and body; `serif` only for pull-quotes
- [ ] One headline size (14 or 16pt) throughout
- [ ] Readable text is `ink` on white or white on `ink` (no small text on a vivid accent)
- [ ] Palette respected: no off-brand colours introduced
- [ ] Generous whitespace; nothing feels cramped
- [ ] Figures use only design variables, match the slide background, and carry no drop shadows
