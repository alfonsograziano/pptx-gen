# 0002 — Native shapes are the output contract

**Status:** Accepted

## Context

Decks built here are opened in Google Slides as often as in PowerPoint, and Google Slides cannot edit or recolor an embedded image. A PNG is a picture; an embedded SVG is worse, showing as a broken-image placeholder. Whatever arrives as an image arrives frozen — the recipient cannot fix a typo in it, cannot recolor it for a different client, cannot move one box six inches left.

It is much easier to draw a diagram as an image and drop it on a slide. That ease is the trap: the deck looks right in review and is dead on arrival for whoever has to edit it next.

## Decision

Everything the engine emits is a shape, a line or a text run. Diagrams, flows, arrows, tables, code panels, timelines and icons are all built from native objects.

Icons are the sharp end of this. `svg-path.ts` parses SVG path data and flattens curves into polylines so an icon can be emitted as native custom geometry — editable and recolorable in both applications — rather than placed as a picture. `helpers.addIcon` and `helpers.addVectorIcon` go through that converter; `helpers.addSvgDiagram` embeds an image and is reserved for genuinely complex vector art that has no shape-based expression.

The test an author applies is in [`custom-template-instructions.md`](../../custom-template-instructions.md): *name the five things a viewer would want to click and change.* If they can be named, build it natively.

## Consequences

- Recipients can actually edit what they receive, in either application, which is the entire point of generating a real `.pptx` instead of a PDF.
- The drawing helpers in `custom-slide-helpers.ts` have to cover a lot of ground, because "just rasterize it" is not available as an escape.
- Curves are flattened to line segments in the icon converter. At icon scale this is invisible; it is a deliberate trade for a small, robust converter.
- **Figures are the one exception, and the bar is high.** Some content has no native form at all: a product UI, a chart with real data, a rendered document, photographic illustration. For those, the author writes HTML and the engine rasterizes it — see [`figure-instructions.md`](../../figure-instructions.md). A figure is not editable in Google Slides, and that cost is the whole reason the exception is kept narrow. It is not a loophole for diagrams, icons, cards or timelines, all of which have native forms and keep them.
