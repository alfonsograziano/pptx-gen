---
name: pptx-deck
description: Generate a client-ready PowerPoint deck from the filesystem template library in this repo. Use when the user gives context, a brief, a narrative, notes, or raw content and asks to create, build, compile, regenerate, or QA a .pptx deck from the imported slide templates. Also use when asked to choose slide templates, write a deterministic build.ts script, run the local generator, inspect rendered screenshots, fix build failures, or return the final PPTX and report.
---

# Build a PPTX deck

## Purpose

Create a PowerPoint deck by writing a deterministic TypeScript project script
that uses the local generator. The generator clones real template slides, fills
their editable text fields, applies explicit overrides, can add project-local
custom slides when no template fits, renders screenshots (if LibreOffice is
installed), and writes a build report.

You can clone template slides, design slides from scratch, or mix both in one
deck. Cloning is fastest when a suitable template exists and you want to preserve
an existing look; designing from scratch is the right path when no template fits,
when the deck is bespoke, or when there is no template library yet. Prefer cloning
only when a genuinely suitable template exists.

## Fixed paths (relative to the repo root)

- Static engine: `src/`
- Template library: `templates/`
- Icon library: `assets/icons/`
- Design system: `design.md` (and its code source `src/design.ts`)
- Custom slide instructions: `custom-template-instructions.md`
- Figure instructions: `figure-instructions.md`
- Generated decks: `projects/<deck-id>/`

## Read the design system first

Before writing `build.ts`, choosing templates, or running the generator, read
`design.md`. It defines the colours, fonts, the sentence-case convention, the
slide grid, and the layout conventions every deck follows. If the user has
customised the design, this is where their choices live.

Use Node 20 or newer:

```bash
node --version
```

## Workflow

### 1. Read the brief

Extract: audience and objective, context, the narrative arc, must-have slides,
claims and numbers and their sources, the visuals the deck needs, assets or
images that must appear, and the output name. If the brief is thin, make
conservative choices and leave clear TODOs in `brief.md`. Do not invent facts or
inflate numbers.

### 1a. Plan the visuals

Before choosing templates or writing any code, decide what each slide *shows*,
and write a `## Visuals` section into `brief.md`. One line per visual: the slide
it belongs to, what it shows, and which of three routes it takes.

```markdown
## Visuals

- Slide 3 — Target architecture → native (boxes and arrows, `custom.ts`)
- Slide 5 — The alerts console as the operator sees it → figure (UI mockup, 1000x560)
- Slide 7 — Adoption by quarter → figure (bar chart, 1280x720)
- Slide 9 — Team photo from the summit → placeholder (only the user has it)
```

The three routes, in order of preference:

1. **Native shapes** — anything made of parts a viewer might move or recolor:
   diagrams, flows, cards, timelines, icons, simple labelled charts. This covers
   most visuals. See 4a.
2. **Figure** — content with no parts: a product or app UI, a chart with a
   continuous axis or many series, a rendered document or terminal. You author
   the HTML and the engine renders it. See 4d.
3. **Image placeholder** — real-world imagery only the user has: a photo, a real
   product screenshot, a client logo. See 4b.

A figure is something you can *draw*; a placeholder is something you must be
*given*.

If the brief also asks for variants of a slide (see 4c), plan the visual once for
the concept and note which variants carry it — a figure is rendered once and can
be placed in several variants.

Rules for this step:

- **Propose a figure wherever one genuinely strengthens the slide**, and say why
  in one clause. A deck explaining a product is much better with the product in
  it; a deck of claims and numbers is better with the numbers drawn.
- **Respect the caps**: at most one figure per slide, figures on no more than one
  slide in five, and never the slide's title or main message inside a figure.
- **The user steers.** The `## Visuals` list is a proposal. Offer it before
  building so they can approve, trim, or extend it — "no figures" is a perfectly
  good answer, and it falls back to native shapes and placeholders. Do not let a
  deck acquire pictures nobody asked for.

### 2. Inspect available templates

List template folders and read their guidance:

```bash
find templates -maxdepth 2 \( -name template.yml -o -name description.md -o -name fields.yml \) | sort
```

For each candidate slide, read:

- `description.md` for when to use the slide and its field list.
- `template.yml` for tags, fonts, and variables.
- `fields.yml` for the exact editable field ids.
- `screenshots/slide-01.png` when it exists and visual choice matters.

Prefer fewer strong slides over many weak ones.

### 3. Create the deck project

```text
projects/<deck-id>/
  build.ts
  brief.md
  figures/        # authored figure HTML, if the deck has any (see 4d)
  inputs/
  output/
```

Keep all deck-specific files inside this folder. `figures/` is source and is
kept; the engine writes the rendered PNGs, and a copy of each figure's HTML, to
`output/figures/`.

### 4. Write `build.ts`

Use the static API. Do not edit `src/` for normal deck generation.

```ts
import { Presentation, md } from "../../src/index.js";

const deck = new Presentation({
  title: "Deck title",
  templateLibrary: "templates",
  projectDir: "projects/<deck-id>",
});

deck.addSlideFromTemplate({
  templateName: "title-cover",
  variables: {
    "overline-label": "Proposal",
    "your-presentation-title-goes-here": "A practical path to production",
    "a-short-subtitle-that-sets-up-the-st": md("From experiments to **safe** delivery."),
  },
});

await deck.render({
  output: "output/deck.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots",
});
```

Rules:

- Use only template names that exist in `templates/`.
- Fill every required variable, using the field ids from `fields.yml`.
- Preserve the original layout unless an override is needed.
- Use `md(...)` for markdown input. The engine converts markdown to styled plain
  text (it does not yet build mixed runs).
- Use overrides only when the template cannot fit the content naturally.
- Keep the script deterministic: no LLM calls, network calls, changing dates, or
  random values during render.

### 4a. Design slides from scratch

For any slide with no fitting template (covers, section breaks, diagrams, flows,
code panels, tables, timelines, or a bespoke look), read
`custom-template-instructions.md` first, then create
`projects/<deck-id>/custom.ts` with named layout functions and call them from
`build.ts`:

```ts
import { Presentation } from "../../src/index.js";
import { architectureFlowSlide } from "./custom.js";

const deck = new Presentation({
  title: "Deck title",
  templateLibrary: "templates",
  projectDir: "projects/<deck-id>",
});

deck.addCustomSlide(architectureFlowSlide({ title: "Target architecture_", nodes: [], arrows: [] }));
```

Rules:

- Prefer cloned templates when a good match exists.
- Use custom slides for diagrams, arrows, flows, code panels, tables, and
  timelines that would need too many overrides.
- **Build everything from native shapes, lines, text, and vector icons** so it
  stays editable and recolorable in PowerPoint and Google Slides. Use `addShape`
  / `addCard` / `addArrow` / `addConnector` / `addText`, and `addIcon` /
  `addVectorIcon` for icons. Do not render diagrams or icons as images. Content
  with no native form at all — an app screen, a chart, a rendered document — is a
  figure instead (4d). Reserve `addSvgDiagram` (embeds a non-editable image) for
  complex art that cannot be expressed as shapes, and note it when you use it.
- Keep custom slide content deterministic and local.
- **Never pass, compute, or hardcode a page number.** `helpers.addFooter(slide)`
  takes no number, and no slide function should accept a `pageNum`. The build
  writes a live PowerPoint slide-number field, so footers stay right when the
  deck is reordered here or in PowerPoint. The same goes for cloned templates:
  their page-number field is tagged `role: page-number` and is not a variable to
  fill.
- Inspect screenshots for every custom slide when available.

### 4b. Image placeholders

For real-world imagery that only the user has — a team photo, a real product
screenshot, a client logo — do not leave a blank gap and do not fake the picture.
Drop a **placeholder**: a grey box with a centered italic caption saying exactly
what the image should show, so the user can supply the real asset later.

If you could draw the thing yourself, it is a figure (4d), not a placeholder.

Use the helper on any custom slide:

```ts
helpers.addImagePlaceholder(slide, {
  x: 5.2, y: 1.6, w: 4.0, h: 3.0,
  caption: "Photo of the team on stage at re:Invent 2024",
});
```

On a **cloned template** slide there is no placeholder helper: use the
`addFigure` override, which draws the captioned grey box itself when the figure
cannot be rendered, or prefer a custom slide when the layout is image-led.

Rules:

- **Do not overuse it.** Reach for a placeholder only when an image clearly adds
  value; most slides need none. Never use it to pad a thin slide.
- **Build diagrams, flows, icons, and cards natively instead** (see 4a), and draw
  app screens, charts, and rendered artifacts as figures (see 4d). A placeholder
  is only for imagery the tool cannot produce at all.
- **Write a specific caption.** "Photo of the team on stage at re:Invent 2024",
  not "image here". The caption is the brief for whoever supplies the asset.
- **The user steers this.** If they ask for more or fewer images, or to turn
  placeholders off entirely, follow that. When in doubt, prefer fewer.

### 4c. Slide variants

Sometimes the brief asks for **options**: several takes on the same slide so the
user can pick a layout. Treat any request for *variants / options / alternatives
/ "a few ways" / "let me choose"* on a named slide or concept as a variant
request — for example *"give me 3 variants of the agenda slide"*.

Variants of a concept are built as consecutive slides, so a deck of two concepts
where the first has three variants renders **four** slides.

**How many.** Two to four per concept; use three when the brief says "a few" or
gives no number. Never more than four — past that the reviewer skims instead of
comparing. If the brief wants variants on more than three concepts, build them
for the highest-value ones and leave the rest as TODOs in `brief.md`. If it asks
for variants of the whole deck, narrow it to specific slides and say so: a fully
duplicated deck is not reviewable.

**Each variant is a genuinely different layout.** Write one named layout function
per variant in `custom.ts` — a timeline, a card grid, a numbered list, a
two-column split with an image placeholder. These do **not** count as variants: a
changed font size, a different accent colour, reordered bullets, the same layout
on a different background, or one function called with different props. If two
variants would share a layout function, they are not variants — redesign one.

**Rewrite the text for each variant.** The concept is fixed — same message, same
facts, same numbers — but the wording is not. A card grid wants three-to-five
word labels; a lead-plus-bullets layout wants a full-sentence lead; a timeline
wants date-prefixed fragments. Copying identical text into every variant produces
layouts that fit badly. Never change facts, claims, or numbers between variants;
only phrasing and density.

**Group and name them.** Tag every variant of a concept with the same `group` —
a short kebab-case name for the concept, not for the layout:

```ts
deck.addCustomSlide(agendaTimelineSlide({ group: "agenda", ... }));
deck.addCustomSlide(agendaCardsSlide({ group: "agenda", ... }));
deck.addSlideFromTemplate({ templateName: "content-lead-bullets", group: "agenda", variables: { ... } });
```

Each factory in `custom.ts` passes `group` straight through to
`new CustomSlide({ name, group, draw })`. Template and custom variants mix freely
in one group.

Name each `CustomSlide` `<concept>-<layout>` — `agenda-timeline`, `agenda-cards`,
`agenda-numbered` — not `agenda-v1`. The name is what the reviewer reads in the
build report, so it should say what the layout *is*.

**Keep a group consecutive.** Add all variants of a concept one after another,
then move to the next concept. The engine emits a `variant-group-split` warning
if a group's slides are not contiguous; treat it as a bug in `build.ts`.

**Page numbers take care of themselves.** Footers carry live slide-number fields,
so a variant build numbers straight through its inflated deck, and the numbers
renumber on their own once the losing variants are removed. Never pass or
compute a page number to compensate for variants.
### 4d. Figures (authored HTML, rendered at build)

A **figure** is an HTML file you write, rendered by headless Chrome at build time
and placed on the slide as a picture. It implements a line from the `## Visuals`
list in `brief.md` (see 1a).

Use it only for content that has no native form at all: product and app UI
mockups, charts with a continuous axis or many series, rendered text artifacts
(an email, a document, a terminal session), and complex illustrative art.

A figure is **not editable in Google Slides**. That is the whole cost, and it is
why the bar is high. Apply the parts test: name the five things a viewer would
want to click and change. If you can name them, build it natively instead.
Diagrams, icons, cards, and timelines are always native.

**Read `figure-instructions.md` before writing any figure markup.** It defines
the brand CSS variables the engine injects and the rules the HTML must follow.

Put the HTML in `projects/<deck-id>/figures/`. On a custom slide:

```ts
const CONSOLE: Figure = {
  id: "alerts-console",
  htmlFile: "figures/alerts-console.html",
  caption: "The alerts console, with the checkout latency alert firing",
  viewport: { width: 1000, height: 560 },
};

await helpers.addFigure(slide, CONSOLE, { x: 5.15, y: 1.62, w: 4.3, h: 2.408 });
```

On a cloned template slide:

```ts
overrides: [
  { op: "addFigure", id: "alerts-console", figure: CONSOLE, x: 5.15, y: 1.62, w: 4.3, h: 2.408 },
]
```

To fill a picture box a template already has, use `replaceFigure` and read that
field's `w` and `h` from `fields.yml` first, then choose a viewport with the same
ratio so the figure lands exactly.

Rules:

- **One figure per slide, and figures on at most one slide in five.**
- **Never put the slide's title or main message inside a figure.** That text must
  be real text on the slide: searchable, translatable, readable by a screen
  reader, and rendered with the deck's embedded fonts. A figure bakes in whatever
  font the build machine had.
- **Match the viewport's ratio to the slide box**, or pass `viewport: "box"`. A
  mismatch is shrunk to fit and leaves a gap; the report says so.
- **Write a real caption.** It is the text a reader sees in the figure's place
  when it cannot be rendered.
- A missing browser is not a build failure: the deck is produced with captioned
  grey boxes and a warning. A missing `.html` file *is* a failure.
- Run `npm run install-fonts` before building a deck with figures, or they render
  in a substitute font and will visibly not match the slides.

### 5. Run and self-heal

From the repo root:

```bash
npm run build
npm run cli -- build --script projects/<deck-id>/build.ts
```

If the build fails: read the error, fix `build.ts` or the project inputs, and
rerun. Repeat until it succeeds or there is a real engine bug. Do not patch
`src/` unless the failure is clearly a reusable engine bug.

### 6. Review output

Inspect `output/deck.pptx`, `output/report.md`, `output/screenshots/*.png`
(present only if LibreOffice is installed), and `output/figures/` when the deck
has figures. Warnings are allowed; report them. Invalid override targets, missing
templates, missing required fields, invalid asset paths, a missing figure source
file, and a corrupt PPTX are hard failures.

Check the report's `## Figures` section. Any figure listed as `placeholder` did
not render — say which ones and why in your final response, because the deck
shipped a grey box where a picture was meant to be.

The report's `## Slides` section lists every slide in deck order and groups any
variants together with their screenshot filenames. If the deck has variant
groups, check that each group's variants really do look different, then point the
user at those screenshots and ask which one they want.

## Override operations

`delete`, `hide`, `move`, `resize`, `styleText`, `addText`, `addSvg`, `addIcon`,
`addImage`, `addFigure`, `replaceImage`, `replaceFigure`. Use layout overrides
sparingly; if a slide needs many, pick a different template.

```ts
deck.addSlideFromTemplate({
  templateName: "content-lead-bullets",
  variables: {
    "section-header": "Why this works",
    "first-supporting-point-that-backs-up": md("Evaluation gates\nHuman approval\nCost controls"),
  },
  overrides: [
    { op: "hide", target: "a-lead-statement-that-frames-the-thr" },
    { op: "styleText", target: "first-supporting-point-that-backs-up", fontSize: 14 },
  ],
});
```

## Final response

Return the absolute path to the final `.pptx`, the report, and the screenshots
folder (plus `output/figures/` if the deck has figures); any warnings or
limitations, including every figure that fell back to a placeholder and why; and
any facts the user must review before using the deck externally.

If the deck contains variants, list each group with its variants and their
screenshot filenames, say what distinguishes each layout, and ask the user to
pick one per group. Once they choose, delete the losing `addCustomSlide` /
`addSlideFromTemplate` calls and their now-unused layout functions in
`custom.ts`, drop the `group` field from the survivor, and rebuild. Page numbers
renumber themselves, so nothing else needs touching.
