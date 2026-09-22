import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { readFile, readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const SRC = path.dirname(fileURLToPath(import.meta.url));

/**
 * pptxgenjs must never be reached by a static or dynamic `import`.
 *
 * It publishes `dist/pptxgen.es.js` under its `import` condition: ESM source in
 * a `.js` file, in a package with no `"type": "module"`. Node loads that as
 * CommonJS and falls back to `require(esm)`, which is illegal inside an import
 * cycle — and `bin/pptx-gen.mjs` registers the tsx loader in-process, so the
 * engine is always loaded from inside one. On Node 23 and 24 that killed the
 * CLI before it built anything. `src/custom-slide.ts` takes the `require`
 * condition instead, which resolves to a genuine CommonJS build.
 *
 * This is a source check rather than a build, because the runtimes that reject
 * the cycle are Node 23 and 24 specifically: a suite run on Node 25 cannot
 * reproduce the crash, but it can keep the import shape that causes it from
 * coming back.
 */
test("no module reaches pptxgenjs through import", async () => {
  const files = (await readdir(SRC)).filter((file) => file.endsWith(".ts"));
  const offenders: string[] = [];

  for (const file of files) {
    const source = await readFile(path.join(SRC, file), "utf8");
    // Strip comments first: this very file, and the note in custom-slide.ts,
    // both spell out the import that must not exist.
    const code = source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
    if (/\bimport\b[^;\n]*["']pptxgenjs["']/.test(code) || /\bimport\s*\(\s*["']pptxgenjs["']/.test(code)) {
      offenders.push(file);
    }
  }

  assert.deepEqual(offenders, [], `these files import pptxgenjs instead of requiring it: ${offenders.join(", ")}`);
});
