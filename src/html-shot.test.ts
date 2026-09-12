import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { findChrome, readPngSize, screenshotHtml } from "./html-shot.js";

// These exercise the real browser, so they skip themselves where none is
// installed — the same contract `render.ts` has with LibreOffice.

const PAGE = `<main style="display:grid;place-items:center;height:100%">
<div style="background:var(--accent);color:var(--white);padding:20px;font-size:40px">figure</div>
</main>`;

test("headless Chrome renders HTML at exactly the requested pixel size", async (t) => {
  const chrome = await findChrome();
  if (!chrome) return t.skip("No Chrome-family browser installed; skipping rasterization test.");

  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-shot-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const htmlPath = path.join(dir, "page.html");
  const pngPath = path.join(dir, "page.png");
  await writeFile(
    htmlPath,
    `<!doctype html><html><head><style>
    :root { --accent: #3B82F6; --white: #fff; }
    html, body { margin: 0; width: 400px; height: 250px; overflow: hidden; }
  </style></head><body>${PAGE}</body></html>`
  );

  const size = await screenshotHtml({ htmlPath, pngPath, width: 400, height: 250, scale: 2 });

  // The whole reason the renderer reads dimensions back off the PNG instead of
  // trusting the request: if a browser ever stops honouring these flags, this is
  // where it surfaces.
  assert.deepEqual(size, { pxWidth: 800, pxHeight: 500 });

  const bytes = await readFile(pngPath);
  assert.deepEqual(readPngSize(bytes), size);
  assert.ok(bytes.length > 200, "a real render produces more than a stub file");
});

test("a stale PNG is not mistaken for this run's output", async (t) => {
  const chrome = await findChrome();
  if (!chrome) return t.skip("No Chrome-family browser installed; skipping rasterization test.");

  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-shot-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const htmlPath = path.join(dir, "page.html");
  const pngPath = path.join(dir, "page.png");
  await writeFile(htmlPath, "<!doctype html><html><body><p>fresh</p></body></html>");
  // A complete but differently-sized PNG left over from an earlier build.
  await writeFile(
    pngPath,
    Buffer.from(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC",
      "base64"
    )
  );

  const size = await screenshotHtml({ htmlPath, pngPath, width: 320, height: 200, scale: 1 });
  assert.deepEqual(size, { pxWidth: 320, pxHeight: 200 }, "the leftover 1x1 must have been replaced");
});
