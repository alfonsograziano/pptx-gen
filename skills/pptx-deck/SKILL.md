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
claims and numbers and their sources, assets or images that must appear, and the
output name. If the brief is thin, make conservative choices and leave clear
TODOs in `brief.md`. Do not invent facts or inflate numbers.

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
  inputs/
  output/
```

Keep all deck-specific files inside this folder.

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
  `addVectorIcon` for icons. Do not render diagrams or icons as images. Reserve
  `addSvgDiagram` (embeds a non-editable image) for complex art that cannot be
  expressed as shapes, and note it when you use it.
- Keep custom slide content deterministic and local.
- Inspect screenshots for every custom slide when available.

### 4b. Image placeholders

Images carry meaning that text cannot, and they pair well with text — a screenshot,
a product photo, a logo, a chart the tool cannot draw. Whenever the content would
genuinely be stronger with an image **that is not a diagram you can build from
native shapes**, do not leave a blank gap and do not fake the picture. Drop a
**placeholder**: a grey box with a centered italic caption saying exactly what the
image should show, so the user can drop the real asset in later.

Use the helper on any custom slide:

```ts
helpers.addImagePlaceholder(slide, {
  x: 5.2, y: 1.6, w: 4.0, h: 3.0,
  caption: "Screenshot of the dashboard's alerts panel",
});
```

On a **cloned template** slide, use the `addSvg` override with a grey box SVG, or
prefer a custom slide when the layout is image-led.

Rules:

- **Do not overuse it.** Reach for a placeholder only when an image clearly adds
  value; most slides need none. Never use it to pad a thin slide.
- **Build diagrams, flows, icons, and charts natively instead** (see 4a) — a
  placeholder is for real-world imagery the tool cannot draw, not for artwork you
  could express as shapes.
- **Write a specific caption.** "Photo of the team on stage at re:Invent 2024",
  not "image here". The caption is the brief for whoever supplies the asset.
- **The user steers this.** If they ask for more or fewer images, or to turn
  placeholders off entirely, follow that. When in doubt, prefer fewer.

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

Inspect `output/deck.pptx`, `output/report.md`, and `output/screenshots/*.png`
(present only if LibreOffice is installed). Warnings are allowed; report them.
Invalid override targets, missing templates, missing required fields, invalid
asset paths, and a corrupt PPTX are hard failures.

## Override operations

`delete`, `hide`, `move`, `resize`, `styleText`, `addText`, `addSvg`, `addIcon`,
`replaceImage`. Use layout overrides sparingly; if a slide needs many, pick a
different template.

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
folder; any warnings or limitations; and any facts the user must review before
using the deck externally.
