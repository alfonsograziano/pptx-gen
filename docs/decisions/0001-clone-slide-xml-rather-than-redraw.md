# 0001 — A cloned slide is edited, never redrawn

**Status:** Accepted

## Context

The usual way to generate a deck from a template is to read the template's values, then draw a new slide from code. For anything but a plain title-and-bullets layout, that loses. A real branded slide carries circles and dashed connectors, arrowheads with specific geometry, embedded fonts, gradient fills, grouped shapes, cropped images and exact type sizes. Reproducing all of it from a drawing API is a large amount of work that is never quite finished, and every near-miss is visible to whoever owns the brand.

The people using this tool already have decks they like. What they want is the slide they already have, with different words in it.

## Decision

Filling a field or applying an override edits the original slide's XML in place. The engine never redraws a cloned slide.

`PptxPackage` (`src/pptx-package.ts`) opens the template `.pptx` as a zip, and `ooxml.ts` performs surgery on the slide XML: `appendSlideFromPackage` copies the slide and its relationships into the deck being built, `fillSlideText` replaces the text inside existing runs, and `applyOverrides` applies explicit, named edits — move, resize, hide, delete, restyle, add. Everything the author did not ask to change is carried across untouched, because it is literally the same XML.

**A cloned slide that renders differently from its source is a bug.** That is the acceptance test for any change in this area.

## Consequences

- Fidelity is free and does not decay. Features of OOXML the engine knows nothing about survive a build, because nothing in the pipeline has to understand them.
- Editing is explicit. There is no implicit re-layout, so text that no longer fits its box stays not fitting: the author moves or resizes it deliberately. This is the right trade — a surprise re-layout on a client's branded slide is worse than a visible overflow the author can see in the screenshot.
- The engine carries real OOXML complexity: relationship rewriting, content-type registration, embedded-font merging (`mergeEmbeddedFonts`) and slide-number field handling all exist because slides are moved between packages rather than built fresh. `ooxml.ts` is the largest module in the repo and that is intrinsic, not accidental.
- Templates must be single-slide slices that share a support chain — layouts, masters, themes — with the deck they are appended into. `sliceToSingleSlide` at ingest time is what guarantees this.
- Slides that do not exist yet cannot be served this way at all, which is the other half of the system: [0002](0002-native-shapes-are-the-output-contract.md).
