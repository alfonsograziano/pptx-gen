# Design system

This is the default design system for decks built with pptx-gen. It is a clean,
neutral starting point. Everything here maps to constants in
[`src/design.ts`](src/design.ts), which is the code source of truth. Edit both
together (or run the `customize-design` skill) to make every deck match your own
brand.

> **One rule worth keeping: sentence case.** Titles, headings, labels, and
> buttons read best in sentence case: only the first word and proper nouns are
> capitalised ("How we ship faster", not "How We Ship Faster"). It reads as more
> human and is easier to translate. This is a convention, not a hard constraint.

---

## Colours

Colours live in `C` in `src/design.ts` as 6-digit hex without a leading `#`.

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

Font families live in `FONTS` in `src/design.ts`.

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

Layout constants live in `LAYOUT` in `src/design.ts`, in inches.

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
  logo mark bottom-right.

---

## Logos (optional)

Logo file names live in `LOGO_FILES` in `src/design.ts`. Drop PNGs with these
names into the `assets/` folder to have them appear automatically:

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

## Quick checklist

Before shipping a deck, verify:

- [ ] Text is sentence case
- [ ] `sans` for headings and body; `serif` only for pull-quotes
- [ ] One headline size (14 or 16pt) throughout
- [ ] Readable text is `ink` on white or white on `ink` (no small text on a vivid accent)
- [ ] Palette respected: no off-brand colours introduced
- [ ] Generous whitespace; nothing feels cramped
