# 0007 — `any` is confined to the XML layer

**Status:** Accepted

## Context

fast-xml-parser returns a plain object tree: keys are element names, attributes are `@_`-prefixed, and a value is a node, an array of nodes, or a scalar depending on what the document happened to contain. Only the OOXML schema pins that shape down, and writing types for the parts of OOXML this engine touches would be a large, permanently incomplete side project.

The realistic options were to spread `any` across every call site in `ooxml.ts`, or to name the untyped thing once.

## Decision

`XmlNode` in `src/xml.ts` is the one documented escape hatch, with a single `biome-ignore` comment explaining why. Everything else in the codebase is strictly typed, and `tsconfig.json` sets `"strict": true`.

**New `any` anywhere outside `src/xml.ts` is a lint error, and should stay one.**

## Consequences

- The deep index chains that XML surgery requires stay readable — `presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]` — without dragging looseness into the rest of the engine.
- The boundary is visible. One `biome-ignore`, in one file, is easy to audit; a hundred scattered ones are not.
- Everything crossing out of the XML layer is typed: `ooxml.ts` takes and returns `TemplateField`, `SlideOverride`, `BuildWarning` and friends from `types.ts`, so the untyped region ends at the module edge.
- Mistakes inside that region are caught by tests rather than by the compiler, which is why `ooxml.test.ts`, `pptx-package.test.ts` and `integration.test.ts` work against real `.pptx` files instead of constructed fixtures.
