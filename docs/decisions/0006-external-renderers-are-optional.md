# 0006 — External renderers are optional and degrade to warnings

**Status:** Accepted

## Context

Two jobs here need a rendering engine nobody wants to bundle. Screenshots for review need something that can open a `.pptx`, and the free cross-platform answer is LibreOffice. Figures need something that can lay out HTML, and the answer is a Chrome-family browser.

Making either a hard dependency would mean a deck cannot be built on a machine that lacks a 300 MB office suite — for a step that is only a preview. Bundling a browser would mean shipping Puppeteer's download for a feature most decks never use.

## Decision

Both are detected at run time and both are optional. Producing a `.pptx` requires no external binary at all.

`render.ts` probes for `soffice` / `libreoffice`; `html-shot.ts` probes a per-platform list of Chrome, Chromium, Edge and Brave paths, overridable with `$CHROME_PATH` and memoized so detection runs at most once. When one is missing, the build pushes a `BuildWarning` naming what was skipped, why, and how to enable it — and then carries on. A figure with no browser becomes a captioned grey placeholder drawn from the figure's own `caption`, which is why the caption is specified as a sentence written for a reader rather than a label.

**Anything depending on these binaries must degrade to a warning and still produce a deck.**

## Consequences

- A deck builds in CI, in a container, or on a fresh laptop with nothing installed.
- Every consumer of the build has to read `report.warnings` rather than trusting exit status, because a successful build can still have skipped a preview. This is also why `output/report.md` exists and why a reviewing agent is told to read it.
- The degraded paths are real code paths and are tested: `figure-integration.test.ts` covers the placeholder route, and `ShotFn` is injectable precisely so the suite needs no browser.
- Two subtleties come with the renderers. LibreOffice resolves `slidenum` fields but ignores the deck's `firstSlideNum` offset, so the file that gets screenshotted is a throwaway copy with its number fields flattened; the delivered deck keeps live fields. And a figure rendered on a machine without the brand's font silently substitutes, so `checkFonts` warns once per build and points at `npm run install-fonts`.
- Figure output is content-addressed against the detected Chrome version, so upgrading the browser invalidates cached PNGs rather than mixing renderings from two engines in one deck.
