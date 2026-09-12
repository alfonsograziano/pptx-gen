import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";
import { loadTemplate } from "./templates.js";
import { ingestTemplate } from "./ingest.js";
import type { Figure } from "./figure.js";
import { fakeShot, TINY_PNG } from "./test-fixtures.js";

// Figures placed on both kinds of slide, verified in the produced .pptx. The
// rasterizer is stubbed, so these run identically with or without a browser.

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATES = path.resolve(HERE, "..", "templates");
const ASSETS = path.resolve(HERE, "..", "assets");

const EMU_PER_IN = 914400;
const emu = (inches: number) => Math.round(inches * EMU_PER_IN);

const MOCKUP: Figure = {
  id: "console-mockup",
  caption: "The alerts console with a firing alert selected",
  html: "<main><h1>Alerts</h1></main>",
  viewport: { width: 1000, height: 250 }
};

async function slideXml(pptxPath: string, index: number): Promise<string> {
  const pkg = await PptxPackage.load(pptxPath);
  const entries = await getSlideEntries(pkg);
  return pkg.text(`ppt/slides/slide${entries[index].slideNumber}.xml`);
}

test("a figure on a custom slide is placed as a picture, inscribed in its box", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const shot = fakeShot();
  const deck = new Presentation({ title: "Figures", templateLibrary: TEMPLATES, projectDir: dir, assetsDir: ASSETS, shot });

  let placed: { x: number; y: number; w: number; h: number } | undefined;
  deck.addCustomSlide(new CustomSlide({
    name: "product",
    draw: async ({ slide, helpers }) => {
      helpers.addHeader(slide, "What the operator sees");
      // A 4:1 figure in a 4:2 box: it must keep its shape and centre itself.
      placed = await helpers.addFigure(slide, MOCKUP, { x: 5, y: 1.5, w: 4, h: 2 });
    }
  }));

  const report = await deck.render({ output: "deck.pptx", progress: false });

  assert.deepEqual(placed, { x: 5, y: 2, w: 4, h: 1 }, "kept its 4:1 aspect and centred vertically");
  assert.equal(shot.calls, 1);

  assert.equal(report.figures.length, 1);
  assert.equal(report.figures[0].status, "rendered");
  assert.equal(report.figures[0].id, "console-mockup");
  assert.equal(report.figures[0].pxWidth, 2000, "viewport times the device scale factor");

  const xml = await slideXml(path.join(dir, "deck.pptx"), 0);
  assert.match(xml, /<p:pic>/, "the figure lands as a picture");
  assert.match(xml, new RegExp(`<a:ext cx="${emu(4)}" cy="${emu(1)}"/>`), "at the inscribed size");

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const media = pkg.files("ppt/media/").filter((file) => file.endsWith(".png"));
  assert.equal(media.length, 1);
  assert.ok((await pkg.bytes(media[0])).equals(TINY_PNG), "the exact rendered bytes are embedded");
});

test("a figure override drops a picture onto a cloned template slide", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir, assetsDir: ASSETS, shot: fakeShot() });
  deck.addSlideFromTemplate({
    templateName: "title-cover",
    variables: { "overline-label": "Figures" },
    overrides: [
      { op: "addFigure", id: "chart", figure: { ...MOCKUP, id: "chart" }, x: 1, y: 1, w: 4, h: 2 }
    ]
  });

  const report = await deck.render({ output: "deck.pptx", progress: false });
  assert.equal(report.figures[0].status, "rendered");

  const xml = await slideXml(path.join(dir, "deck.pptx"), 0);
  assert.match(xml, /name="chart"/);
  assert.match(xml, new RegExp(`<a:ext cx="${emu(4)}" cy="${emu(1)}"/>`), "inscribed, not stretched");
});

test("without a browser the deck still builds, with captioned placeholders", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir, assetsDir: ASSETS, shot: fakeShot({ fail: true }) });
  deck.addSlideFromTemplate({
    templateName: "title-cover",
    variables: { "overline-label": "Figures" },
    overrides: [{ op: "addFigure", id: "chart", figure: { ...MOCKUP, id: "chart" }, x: 1, y: 1, w: 4, h: 2 }]
  });
  deck.addCustomSlide(new CustomSlide({
    name: "product",
    draw: async ({ slide, helpers }) => {
      await helpers.addFigure(slide, MOCKUP, { x: 5, y: 1.5, w: 4, h: 2 });
    }
  }));

  const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

  assert.equal(report.figures.length, 2);
  assert.ok(report.figures.every((figure) => figure.status === "placeholder"));
  assert.equal(report.warnings.filter((warning) => warning.code === "figure-render-failed").length, 2);

  const cloned = await slideXml(path.join(dir, "deck.pptx"), 0);
  assert.match(cloned, /prstGeom prst="roundRect"/, "a dashed placeholder box stands in");
  assert.match(cloned, /<a:t>The alerts console with a firing alert selected<\/a:t>/, "carrying the caption verbatim");
  assert.doesNotMatch(cloned, /name="chart"[\s\S]*<p:pic>/, "and no picture was inserted");

  // The custom slide falls back through the same placeholder helper.
  const custom = await slideXml(path.join(dir, "deck.pptx"), 1);
  assert.match(custom, /The alerts console with a firing alert selected/);

  const markdown = await readFile(path.join(dir, "report.md"), "utf8");
  assert.match(markdown, /## Figures/);
  assert.match(markdown, /console-mockup: placeholder/);
});

/**
 * Build a one-slide template that contains a real picture, so the replace* paths
 * have something to target. The shipped library has no picture fields, and the
 * geometry rewriting these overrides do is the part most worth pinning down.
 */
async function templateWithPicture(dir: string, box: { w: number; h: number }): Promise<{ root: string; fieldId: string }> {
  const seed = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir, assetsDir: ASSETS });
  const imagePath = path.join(dir, "seed.png");
  await writeFile(imagePath, TINY_PNG);
  seed.addSlideFromTemplate({
    templateName: "title-cover",
    variables: { "overline-label": "Seed" },
    overrides: [{ op: "addImage", id: "photo", path: "seed.png", x: 1, y: 1, w: box.w, h: box.h }]
  });
  await seed.render({ output: "seed.pptx", progress: false });

  const root = path.join(dir, "library");
  await ingestTemplate({ source: path.join(dir, "seed.pptx"), templateRoot: root, templateName: "with-picture", slide: 1 });

  const template = await loadTemplate(root, "with-picture");
  const field = template.fieldsFile.fields.find((candidate) => candidate.type === "image");
  assert.ok(field, "the ingested template should expose its picture as a field");
  return { root, fieldId: field.id };
}

test("replaceFigure re-inscribes the picture box around the figure's own shape", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  // A square hole in the template, and a 4:1 figure going into it.
  const { root, fieldId } = await templateWithPicture(dir, { w: 2, h: 2 });

  const deck = new Presentation({ templateLibrary: root, projectDir: dir, assetsDir: ASSETS, shot: fakeShot() });
  deck.addSlideFromTemplate({
    templateName: "with-picture",
    overrides: [{ op: "replaceFigure", target: fieldId, figure: MOCKUP }]
  });
  await deck.render({ output: "replaced.pptx", progress: false });

  const xml = await slideXml(path.join(dir, "replaced.pptx"), 0);
  assert.match(
    xml,
    new RegExp(`<a:ext cx="${emu(2)}" cy="${emu(0.5)}"/>`),
    "the box shrinks to the figure's 4:1 shape rather than squashing it"
  );
  assert.match(xml, new RegExp(`<a:off x="${emu(1)}" y="${emu(1.75)}"/>`), "and is centred in the original box");
});

test("replaceFigure keeps the template's own image when the figure cannot render", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const { root, fieldId } = await templateWithPicture(dir, { w: 2, h: 2 });
  const before = await slideXml(path.join(dir, "seed.pptx"), 0);
  const originalRel = before.match(/<a:blip\b[^>]*r:embed="([^"]+)"/)?.[1];

  const deck = new Presentation({ templateLibrary: root, projectDir: dir, assetsDir: ASSETS, shot: fakeShot({ fail: true }) });
  deck.addSlideFromTemplate({
    templateName: "with-picture",
    overrides: [{ op: "replaceFigure", target: fieldId, figure: MOCKUP }]
  });

  const report = await deck.render({ output: "kept.pptx", progress: false });
  assert.ok(report.warnings.some((warning) => warning.code === "figure-placeholder-used"));

  const xml = await slideXml(path.join(dir, "kept.pptx"), 0);
  assert.match(xml, new RegExp(`<a:blip\\b[^>]*r:embed="${originalRel}"`), "the original picture is untouched");
  assert.match(xml, new RegExp(`<a:ext cx="${emu(2)}" cy="${emu(2)}"/>`), "and so is its box");
});

test("two custom slides carrying different images keep their own media", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-figure-int-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  // A regression guard: every custom slide is rendered in its own package and
  // each numbers its media from image1, so without collision handling the second
  // slide would silently display the first slide's picture.
  const red = path.join(dir, "red.png");
  const blue = path.join(dir, "blue.png");
  await writeFile(red, TINY_PNG);
  await writeFile(blue, Buffer.concat([TINY_PNG, Buffer.from("distinct-tail")]));

  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir, assetsDir: ASSETS });
  deck.addSlideFromTemplate({ templateName: "title-cover", variables: { "overline-label": "Media" } });
  for (const [index, image] of [red, blue].entries()) {
    deck.addCustomSlide(new CustomSlide({
      name: `picture-${index}`,
      draw: ({ slide }) => {
        slide.addImage({ path: image, x: 1, y: 1, w: 2, h: 2 });
      }
    }));
  }

  await deck.render({ output: "deck.pptx", progress: false });

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const entries = await getSlideEntries(pkg);
  const targets: string[] = [];
  for (const entry of entries.slice(1)) {
    const rels = await pkg.text(`ppt/slides/_rels/slide${entry.slideNumber}.xml.rels`);
    const image = [...rels.matchAll(/Target="(\.\.\/media\/[^"]+)"/g)]
      .map((match) => match[1])
      .find((target) => /\.(png|jpe?g)$/i.test(target));
    if (image) targets.push(image);
  }

  assert.equal(targets.length, 2);
  assert.notEqual(targets[0], targets[1], "each slide must reference its own picture part");

  const bytes = await Promise.all(targets.map((target) => pkg.bytes(path.posix.join("ppt", target.replace("../", "")))));
  assert.ok(!bytes[0].equals(bytes[1]), "and those parts must hold different bytes");
});
