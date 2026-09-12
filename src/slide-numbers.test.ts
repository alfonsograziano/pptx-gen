import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, rm } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { PptxPackage } from "./pptx-package.js";
import { flattenSlideNumberFields, getSlideEntries } from "./ooxml.js";
import type { BuildWarning, SlideOverride } from "./types.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATES = path.resolve(HERE, "..", "templates");

type SlideShape = "cover" | { custom: string } | { template: string; overrides?: SlideOverride[] };

/** What a slide ended up carrying, read back out of the saved package. */
type SlideFacts = {
  label: string;
  /** Text of the slide-number field, or undefined when the slide has none. */
  number?: string;
  hidden: boolean;
  /** True if a plain (non-field) run still holds a hardcoded "2" from the template. */
  frozenLiteral: boolean;
};

function footered(label: string): CustomSlide {
  return new CustomSlide({
    name: label,
    draw({ slide, helpers }) {
      helpers.addHeader(slide, label);
      helpers.addFooter(slide);
    }
  });
}

async function buildDeck(dir: string, name: string, shapes: SlideShape[]): Promise<{
  output: string;
  warnings: BuildWarning[];
}> {
  const deck = new Presentation({ title: name, templateLibrary: TEMPLATES, projectDir: dir });
  for (const [index, shape] of shapes.entries()) {
    if (shape === "cover") {
      deck.addSlideFromTemplate({
        templateName: "title-cover",
        variables: { "your-presentation-title-goes-here": `slide-${index}` }
      });
    } else if ("custom" in shape) {
      deck.addCustomSlide(footered(shape.custom));
    } else {
      deck.addSlideFromTemplate({
        templateName: shape.template,
        variables: { "section-header": `slide-${index}` },
        overrides: shape.overrides
      });
    }
  }
  const output = path.join(dir, `${name}.pptx`);
  const report = await deck.render({ output, screenshots: path.join(dir, `${name}-shots`), progress: false });
  return { output, warnings: report.warnings };
}

async function readSlides(output: string): Promise<SlideFacts[]> {
  const pkg = await PptxPackage.load(output);
  const entries = await getSlideEntries(pkg);
  return Promise.all(entries.map(async (entry) => {
    const xml = await pkg.text(`ppt/slides/slide${entry.slideNumber}.xml`);
    return {
      label: xml.match(/<a:t>(slide-\d+|[A-Z]\w*)<\/a:t>/)?.[1] ?? "?",
      number: xml.match(/type="slidenum">[\s\S]*?<a:t>([^<]*)<\/a:t>/)?.[1],
      hidden: /<p:cNvPr\b[^>]*\bhidden="1"/.test(xml),
      // A run, as opposed to a field, still holding the number the source deck
      // was exported with.
      frozenLiteral: /<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t>2<\/a:t>/.test(xml)
    };
  }));
}

async function firstSlideNum(output: string): Promise<number> {
  const pkg = await PptxPackage.load(output);
  const xml = await pkg.text("ppt/presentation.xml");
  return Number(xml.match(/firstSlideNum="(-?\d+)"/)?.[1] ?? 1);
}

test("page numbers follow deck order, so reordering slides renumbers them", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-slide-numbers-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const shapes: SlideShape[] = ["cover", { custom: "Alpha" }, { custom: "Beta" }];
  const normal = await buildDeck(dir, "normal", shapes);
  // The same deck with two slides swapped and nothing else changed.
  const swapped = await buildDeck(dir, "swapped", ["cover", { custom: "Beta" }, { custom: "Alpha" }]);

  const before = await readSlides(normal.output);
  const after = await readSlides(swapped.output);

  assert.deepEqual(before.map((slide) => [slide.label, slide.number]), [
    ["slide-0", undefined],
    ["Alpha", "1"],
    ["Beta", "2"]
  ]);
  assert.deepEqual(after.map((slide) => [slide.label, slide.number]), [
    ["slide-0", undefined],
    ["Beta", "1"],
    ["Alpha", "2"]
  ]);
});

test("numbering starts at 1 on the first slide that shows a footer", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-slide-numbers-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  // A cover takes position 1 but draws no footer, so the deck counts from 0 to
  // leave the first footered slide reading "1".
  const withCover = await buildDeck(dir, "with-cover", ["cover", { custom: "Alpha" }]);
  assert.equal(await firstSlideNum(withCover.output), 0);

  // With no cover the first slide is already "1", so the attribute stays off.
  const noCover = await buildDeck(dir, "no-cover", [{ custom: "Alpha" }, { custom: "Beta" }]);
  assert.equal(await firstSlideNum(noCover.output), 1);

  assert.equal(
    withCover.warnings.some((warning) => warning.code === "page-numbering-clamped"),
    false,
    "one unnumbered leading slide is within what PowerPoint can offset"
  );
});

test("a template's baked-in page number is replaced, never shipped", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-slide-numbers-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  // content-lead-bullets was ingested from slide 2 of the example deck and
  // carries a literal "2" in its footer.
  const built = await buildDeck(dir, "template", [
    "cover",
    { custom: "Alpha" },
    { template: "content-lead-bullets" }
  ]);
  const slides = await readSlides(built.output);

  assert.equal(slides[2].number, "2", "the template slide reads its real position");
  assert.equal(slides[2].frozenLiteral, false, "the source deck's hardcoded 2 is gone");
});

test("a hidden footer is not what numbering counts from", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-slide-numbers-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const built = await buildDeck(dir, "hidden", [
    "cover",
    { template: "content-lead-bullets", overrides: [{ op: "hide", target: "page-number" }] },
    { custom: "Alpha" }
  ]);
  const slides = await readSlides(built.output);

  assert.equal(slides[1].hidden, true);
  // Two slides show nothing before the first visible footer, which would need
  // PowerPoint to count from -1. It cannot, so the offset clamps and the build
  // says so rather than silently numbering from the wrong place.
  assert.equal(await firstSlideNum(built.output), 0);
  assert.ok(
    built.warnings.some((warning) => warning.code === "page-numbering-clamped"),
    `expected a clamp warning, got: ${built.warnings.map((w) => w.code).join(", ")}`
  );
});

test("flattening replaces each field with the number it resolves to", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-slide-numbers-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  // This is what the screenshot copy goes through: LibreOffice honours slidenum
  // fields but ignores the deck's offset, so the preview is shot from a copy
  // with the fields resolved to plain text.
  const built = await buildDeck(dir, "flatten", ["cover", { custom: "Alpha" }, { custom: "Beta" }]);
  const pkg = await PptxPackage.load(built.output);
  await flattenSlideNumberFields(pkg);

  const entries = await getSlideEntries(pkg);
  const slides = await Promise.all(entries.map((entry) => pkg.text(`ppt/slides/slide${entry.slideNumber}.xml`)));

  for (const xml of slides) {
    assert.ok(!/type="slidenum"/.test(xml), "no field survives flattening");
  }
  assert.ok(/<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t>1<\/a:t>/.test(slides[1]), "first footered slide flattens to 1");
  assert.ok(/<a:r>(?:(?!<\/a:r>)[\s\S])*?<a:t>2<\/a:t>/.test(slides[2]), "the next one flattens to 2");
});
