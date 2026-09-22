# 0009 — pptxgenjs is loaded through `require`, lazily

**Status:** Accepted

## Context

pptxgenjs publishes `dist/pptxgen.es.js` under its `import` condition: ESM source in a file ending `.js`, inside a package with no `"type": "module"`. Node therefore has to load that file as CommonJS and fall back to `require(esm)`. That fallback works, with one exception — it is illegal inside an import cycle, and Node 23 and 24 reject it outright:

```
Cannot require() ES Module .../pptxgen.es.js in a cycle. A cycle involving
require(esm) is not allowed to maintain invariants mandated by the ECMAScript
specification.
```

`bin/pptx-gen.mjs` registers the tsx ESM loader in-process and then imports `src/cli.ts`, so the whole engine is loaded from inside a live cycle. Every `pptx-gen build` died before building anything on those two Node majors, while `npx tsx src/cli.ts build` — the same code with the loader started before the graph — worked. `package.json` declares `engines: { node: ">=20" }`, so this was a broken promise, not an unsupported configuration.

Moving the import from module top level to a lazy `await import("pptxgenjs")` at the point of use does **not** fix it. It moves the crash from startup to the first custom slide: the cycle is around the loader, not around the module graph's shape, so deferring the same `import` changes nothing.

## Decision

`src/custom-slide.ts` loads pptxgenjs with `createRequire(import.meta.url)("pptxgenjs")`, on first use, memoized.

That takes the package's `require` condition, which is `dist/pptxgen.cjs.js` — a genuine CommonJS file. No `require(esm)` happens, so there is no cycle to be inside of, on any Node version.

**No module in `src/` may reach pptxgenjs through a static or dynamic `import`.** `src/custom-slide.test.ts` enforces it by scanning the sources, because the suite runs on one Node version and the two that reject the cycle cannot be reproduced from inside it.

## Consequences

- The CLI works on Node 20 through 25, which is what `engines` says.
- Deferring the load to first use also keeps pptxgenjs out of the process entirely for decks built only from cloned templates. That is a side benefit, not the reason; the reason is the cycle.
- The engine takes pptxgenjs's CJS build rather than its ESM one. The two are built from the same source, and the constructor is unwrapped from `.default` if the bundle put it there, so nothing downstream sees the difference.
- A Node-version regression here cannot be caught by `npm test`. Verifying it means running `pptx-gen build` under `nvm` on 23 and 24, which is what was done for this change.
- If pptxgenjs ever ships `dist/pptxgen.es.js` as `.mjs`, or the package gains `"type": "module"`, this is safe to revisit — but the lazy `require` costs nothing, so there is no reason to.
