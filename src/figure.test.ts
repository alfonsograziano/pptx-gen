import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, readdir, rm, writeFile } from "node:fs/promises";
import {
  buildFigureHtml,
  figureHash,
  fitBox,
  FigureRenderer,
  resolveViewport,
  type Figure,
  type FigureResult
} from "./figure.js";
import { readPngSize } from "./html-shot.js";
import { C, FONTS } from "./design.js";
import type { BuildWarning } from "./types.js";
import { fakeShot, TINY_PNG } from "./test-fixtures.js";

const VIEWPORT = { width: 1280, height: 720 };

/** Narrows a result to the success shape, failing the test with its reason if not. */
function rendered(result: FigureResult): Extract<FigureResult, { pngPath: string }> {
  assert.notEqual(result.status, "failed", result.status === "failed" ? result.reason : "");
  if (result.status === "failed") throw new Error("unreachable");
  return result;
}
const figure = (overrides: Partial<Figure> = {}): Figure => ({
  id: "demo",
  caption: "A demo figure",
  html: "<main>hello</main>",
  ...overrides
});

test("buildFigureHtml exposes every design token as a CSS variable", () => {
  const html = buildFigureHtml(figure(), VIEWPORT, "<main>hello</main>");
  for (const [name, value] of Object.entries(C)) {
    assert.ok(html.includes(`#${value}`), `expected the palette value for '${name}' in the figure CSS`);
  }
  assert.match(html, /--ink: #/);
  assert.match(html, /--accent-2: #/, "camelCase and digits become kebab-case variable names");
  assert.match(html, /--grey-10: #/);
  assert.match(html, /--accent-soft: #/);
  assert.ok(html.includes(`--font-sans: "${FONTS.sans}"`));
  assert.match(html, /--font-sans:[^;]*sans-serif/, "the brand font needs a real fallback stack");
});

test("buildFigureHtml pins the page to the viewport so it cannot overflow", () => {
  const html = buildFigureHtml(figure(), { width: 800, height: 500 }, "<main>hello</main>");
  assert.match(html, /width: 800px/);
  assert.match(html, /height: 500px/);
  assert.match(html, /overflow: hidden/);
  assert.match(html, /font-size: 16px/, "rem units must not depend on browser defaults");
});

test("buildFigureHtml injects into an authored document instead of nesting one", () => {
  const source = "<!doctype html><html><head><title>Mine</title></head><body><p>hi</p></body></html>";
  const html = buildFigureHtml(figure({ html: source }), VIEWPORT, source);
  assert.equal(html.match(/<html/gi)?.length, 1, "must not wrap a full document in another one");
  assert.equal(html.match(/<body/gi)?.length, 1);
  assert.ok(html.includes("<title>Mine</title>"), "the author's own head content survives");
  assert.ok(html.indexOf("--ink:") < html.indexOf("</head>"), "brand CSS belongs in the head");
});

test("buildFigureHtml gives a document without a head one to hold the brand CSS", () => {
  const source = "<html><body><p>hi</p></body></html>";
  const html = buildFigureHtml(figure({ html: source }), VIEWPORT, source);
  assert.match(html, /<head>[\s\S]*--ink:[\s\S]*<\/head>/);
});

test("buildFigureHtml appends author CSS after the brand block so it can override", () => {
  const html = buildFigureHtml(figure({ css: "body { outline: 1px solid red; }" }), VIEWPORT, "<main/>");
  assert.ok(html.indexOf("--ink:") < html.indexOf("outline: 1px solid red"));
});

test("figureHash is stable, and changes with anything that changes the pixels", () => {
  const base = figureHash("<p>a</p>", VIEWPORT, 2, false, "152");
  assert.equal(base, figureHash("<p>a</p>", VIEWPORT, 2, false, "152"));
  assert.notEqual(base, figureHash("<p>b</p>", VIEWPORT, 2, false, "152"));
  assert.notEqual(base, figureHash("<p>a</p>", { width: 1281, height: 720 }, 2, false, "152"));
  assert.notEqual(base, figureHash("<p>a</p>", VIEWPORT, 3, false, "152"));
  assert.notEqual(base, figureHash("<p>a</p>", VIEWPORT, 2, true, "152"));
  assert.notEqual(base, figureHash("<p>a</p>", VIEWPORT, 2, false, "153"), "a browser upgrade must invalidate");
});

test("resolveViewport defaults, honours an explicit size, and derives one from the box", () => {
  assert.deepEqual(resolveViewport(figure()), { width: 1280, height: 720 });
  assert.deepEqual(resolveViewport(figure({ viewport: { width: 640, height: 640 } })), { width: 640, height: 640 });
  assert.deepEqual(resolveViewport(figure({ viewport: "box" }), { w: 4, h: 2 }), { width: 512, height: 256 });
  assert.deepEqual(resolveViewport(figure({ viewport: "box" })), { width: 1280, height: 720 }, "no box, no derivation");
});

test("fitBox inscribes the figure in the box without distorting it", () => {
  const box = { x: 1, y: 1, w: 4, h: 2 };

  const wide = fitBox(box, 1000, 250, "contain");
  assert.equal(wide.w, 4, "a figure wider than the box is limited by width");
  assert.equal(wide.h, 1);
  assert.equal(wide.y, 1.5, "and is centred vertically");
  assert.equal(wide.x, 1);

  const tall = fitBox(box, 250, 1000, "contain");
  assert.equal(tall.h, 2, "a figure taller than the box is limited by height");
  assert.equal(tall.w, 0.5);
  assert.equal(tall.x, 2.75, "and is centred horizontally");

  assert.deepEqual(fitBox(box, 800, 400, "contain"), box, "a matching aspect is left exactly alone");
  assert.deepEqual(fitBox(box, 1000, 250, "stretch"), box, "stretch fills the box regardless");
});

test("readPngSize reads the real dimensions out of the IHDR header", () => {
  assert.deepEqual(readPngSize(TINY_PNG), { pxWidth: 1, pxHeight: 1 });
  assert.throws(() => readPngSize(Buffer.from("not a png at all, truly")), /Not a PNG/);
});

async function renderer(t: { after: (fn: () => unknown) => void }, options: { fail?: boolean } = {}) {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-"));
  t.after(() => rm(dir, { recursive: true, force: true }));
  const warnings: BuildWarning[] = [];
  const shot = fakeShot(options);
  const figures = new FigureRenderer({ outputDir: dir, projectDir: dir, warnings, shot });
  return { dir, warnings, shot, figures };
}

test("a figure is rendered once and then served from cache", async (t) => {
  const { dir, shot, figures } = await renderer(t);

  const first = await figures.render(figure());
  assert.equal(first.status, "rendered");
  assert.equal(shot.calls, 1);
  assert.equal(path.dirname(first.htmlPath), dir);

  // A fresh renderer, so the in-memory memo cannot be what answers this.
  const second = new FigureRenderer({ outputDir: dir, projectDir: dir, warnings: [], shot });
  const cached = await second.render(figure());
  assert.equal(cached.status, "cached");
  assert.equal(shot.calls, 1, "identical input must not re-rasterize");
  assert.equal(rendered(cached).pngPath, rendered(first).pngPath);

  const changed = await second.render(figure({ html: "<main>different</main>" }));
  assert.equal(changed.status, "rendered");
  assert.equal(shot.calls, 2, "changed markup must re-rasterize");
});

test("the same figure used twice in one build rasterizes once", async (t) => {
  const { shot, figures } = await renderer(t);
  await Promise.all([figures.render(figure()), figures.render(figure())]);
  assert.equal(shot.calls, 1);
});

test("the rendered PNG reports the pixels it actually has", async (t) => {
  const { figures } = await renderer(t);
  const result = rendered(await figures.render(figure({ viewport: { width: 400, height: 300 }, scale: 2 })));
  assert.equal(result.status, "rendered");
  assert.equal(result.pxWidth, 800);
  assert.equal(result.pxHeight, 600);
});

test("a rasterizer failure degrades to a placeholder and keeps the HTML on disk", async (t) => {
  const { warnings, figures } = await renderer(t, { fail: true });

  const result = await figures.render(figure());
  assert.equal(result.status, "failed");

  const html = await import("node:fs/promises").then((fs) => fs.readFile(result.htmlPath, "utf8"));
  assert.match(html, /--ink:/, "the generated HTML is still written so it can be opened and debugged");

  const failure = warnings.filter((warning) => warning.code === "figure-render-failed");
  assert.equal(failure.length, 1);
  assert.match(failure[0].message, /demo/);
  assert.equal(figures.results()[0].status, "placeholder");
});

test("a viewport that does not match its slide box is reported", async (t) => {
  const { warnings, figures } = await renderer(t);
  await figures.render(figure({ viewport: { width: 1280, height: 720 } }), { box: { w: 3, h: 3 } });
  const mismatch = warnings.filter((warning) => warning.code === "figure-aspect-mismatch");
  assert.equal(mismatch.length, 1);
  assert.match(mismatch[0].message, /1\.78:1.*1\.00:1/);
});

test("a viewport matching its box, and viewport 'box', are not reported", async (t) => {
  const { warnings, figures } = await renderer(t);
  await figures.render(figure({ id: "a", viewport: { width: 1000, height: 500 } }), { box: { w: 4, h: 2 } });
  await figures.render(figure({ id: "b", viewport: "box" }), { box: { w: 3, h: 3 } });
  assert.equal(warnings.filter((warning) => warning.code === "figure-aspect-mismatch").length, 0);
});

test("prune removes superseded figure files and keeps the current ones", async (t) => {
  const { dir, figures } = await renderer(t);
  const result = rendered(await figures.render(figure()));

  await writeFile(path.join(dir, "demo.0123456789ab.png"), TINY_PNG);
  await writeFile(path.join(dir, "removed-figure.html"), "<p>from an older build</p>");

  await figures.prune();

  const remaining = (await readdir(dir)).sort();
  assert.deepEqual(remaining, [path.basename(result.pngPath), "demo.html"].sort());
});

test("a figure with no source, or with two, is an authoring error", async (t) => {
  const { figures } = await renderer(t);
  await assert.rejects(() => figures.render({ id: "empty", caption: "nothing" }), /has no source/);
  await assert.rejects(
    () => figures.render({ id: "both", caption: "two", html: "<p/>", htmlFile: "x.html" }),
    /use one/
  );
  await assert.rejects(
    () => figures.render({ id: "missing", caption: "gone", htmlFile: "nope.html" }),
    /was not found/
  );
});

test("htmlFile is read relative to the deck project", async (t) => {
  const { dir, figures } = await renderer(t);
  await writeFile(path.join(dir, "panel.html"), "<main>from a file</main>");
  const result = await figures.render(figure({ htmlFile: "panel.html", html: undefined }));
  assert.equal(result.status, "rendered");
  const html = await import("node:fs/promises").then((fs) => fs.readFile(result.htmlPath, "utf8"));
  assert.match(html, /from a file/);
});
