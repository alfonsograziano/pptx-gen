# 0004 — Design values are data, layered over defaults

**Status:** Accepted

## Context

Every scratch-built slide needs a palette, a type scale, a grid and a set of logos. If those values live in code, a brand becomes a fork, and an engine update means a merge conflict against someone's colours. If a brand file has to be complete, then adding a token to the engine silently breaks every brand file written before it existed.

## Decision

The engine's defaults live in `src/design.ts`. A workspace's brand lives in its `design.yml`, which is a **partial patch**: anything it omits falls back to the default. `design-loader.ts` reads and validates it; `design-tokens.ts` holds the vocabulary — which tokens exist and their types — and imports nothing, so the values and the loader never depend on each other.

`applyDesign` mutates `C`, `FONTS`, `LAYOUT` and `LOGO_FILES` **in place**. That is deliberate: a deck script writes `const { LM } = LAYOUT` at module top level, and in-place mutation is what lets that capture the brand's value rather than the default.

Adding a token means adding a default *and* letting a partial override fall back to it, so a `design.yml` written months ago keeps working.

## Consequences

- Changing one value in `design.yml` restyles every scratch-built slide at once, with no code change.
- Engine updates are safe for brands. New tokens appear with sensible defaults; old files stay valid.
- In-place mutation creates a hard ordering rule. The design must be applied before any deck script is imported, and never re-applied afterwards, because a script's top-level destructuring captures whatever it sees at import time. `src/cli.ts` carries this note at the two places it matters, and `design-autoload.test.ts` is the regression test for the whole scheme.
- Colour parsing is fussier than it looks. `design-loader.ts` parses the YAML twice — once with the `failsafe` schema — because the core schema turns a digits-only colour into a number lossily: `000000` becomes `0`, `001122` becomes `1122`, `0x1122` becomes `4386`. Colours, font names and logo file names are read from the literal parse; layout numbers come from the typed one.
- The token set is a vocabulary, not a free-form map. `accentOnDark`, `paperSoft` and `inkSoft` exist because one accent cannot clear the contrast bar on both light and dark grounds, and a raised card needs a tone distinct from the page in both directions.
