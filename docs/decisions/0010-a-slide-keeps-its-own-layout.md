# 0010 — A cloned slide keeps its own layout, matched by content

**Status:** Accepted

## Context

A deck is built on the package of its first template slide, and every later slide is appended into it (`appendSlideFromPackage` in `src/ooxml.ts`). The slide XML only says which layout it uses, by path: `../slideLayouts/slideLayout51.xml`. The layout, its master and the master's theme hold the background, the brand marks, the placeholder styles and the theme fonts. Much of what a slide looks like lives there, not on the slide.

The append assumed every template shared one support chain, path for path, with the base. That holds for slices of one source deck. It fails as soon as a workspace ingests from two decks, which real workspaces do. It fails two ways:

- **The path is missing.** The slide points at a layout the base lacks. The build said `missing-slide-dependency` and shipped a package with a dangling relationship, which LibreOffice rendered on the wrong background with a stray wordmark.
- **The path names a different layout.** The base has its own `slideLayout23.xml`, and the slide is bound to it. No warning at all. The slide takes the base's theme fonts and reflows, which breaks [decision 0001](0001-clone-slide-xml-rather-than-redraw.md)'s promise that a cloned slide keeps its pixels.

## Decision

**A layout is matched on content, never on path.** `importSlideLayout` fingerprints the source layout: its XML, the bytes of its pictures, and recursively its master and theme. A master's own list of layouts is left out, since that list does not change how any one layout renders. If the target has a layout with the same fingerprint, the slide is pointed at it, and the same path is tried first because slices of one deck share their parts. Otherwise the layout is copied in under a fresh name. Its master is reused if the target has an equal, and copied otherwise, along with its own copy of the theme.

An imported master starts with no layouts, and each layout joins it only when a slide needs it. A 200-layout source deck contributes one layout, not 200. Master and layout ids are allocated from one shared space, unique and no lower than 2^31, which PowerPoint requires.

Custom slides are the exception: `presentation.ts` passes `importLayout: false`, and a custom slide stays on the base's first layout as it always has. Moving it to pptxgenjs's own blank master would change the theme under every existing custom slide.

## Consequences

- A template from any deck can follow any other. The workspace-level rule "treat `missing-slide-dependency` as a failure, and avoid these fifteen templates" is no longer needed for layouts.
- A deck that mixes source decks carries one extra master, theme and layout per foreign deck it draws on. The output is larger by those parts, usually kilobytes, plus any pictures those masters hold.
- Appending costs a hash of the layout chain on both sides for every template slide, including same-deck slices. It is per-call memoised and cheap next to rendering, but it is not free.
- Equality is strict. Two layouts that differ only in a relationship id count as different, and the layout is imported rather than shared. That wastes a few bytes and is never wrong.
- The slide size is still the base deck's. A template ingested from a deck with a different slide size is not rescaled, and that remains the caller's problem.
