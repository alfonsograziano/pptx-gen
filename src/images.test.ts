import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdir, mkdtemp, readFile, realpath, rm, writeFile } from "node:fs/promises";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";
import { ingestTemplate } from "./ingest.js";
import { loadTemplate } from "./templates.js";
import { coverCrop, imageSize } from "./images.js";
import { BUNDLED_ASSETS, makePng, STARTER_TEMPLATES, TINY_PNG } from "./test-fixtures.js";

// Pictures from the brief's ## Images block, placed on both kinds of slide and
// checked in the produced .pptx. Nothing here generates anything: the files are
// written straight into inputs/, which is all the build ever looks at.

const EMU_PER_IN = 914400;
const emu = (inches: number) => Math.round(inches * EMU_PER_IN);

const BRIEF = `# Deck

## Images

\`\`\`yaml
- id: hero
  slide: 1
  description: A calm harbour at dawn, boats at rest
  variants: 2
  pick: 2
- id: not-yet
  description: The team on stage at the summit
\`\`\`
`;

async function deckProject(t: { after: (fn: () => unknown) => void }): Promise<string> {
  const dir = await realpath(await mkdtemp(path.join(os.tmpdir(), "pptx-images-")));
  t.after(() => rm(dir, { recursive: true, force: true }));
  await writeFile(path.join(dir, "brief.md"), BRIEF, "utf8");
  await mkdir(path.join(dir, "inputs"));
  // Variant 2 is the pick; a 4:1 picture, so every box below needs cropping.
  await writeFile(path.join(dir, "inputs", "hero-1.png"), makePng(20, 20));
  await writeFile(path.join(dir, "inputs", "hero-2.png"), makePng(400, 100));
  return dir;
}

async function slideXml(pptxPath: string, index: number): Promise<string> {
  const pkg = await PptxPackage.load(pptxPath);
  const entries = await getSlideEntries(pkg);
  return pkg.text(`ppt/slides/slide${entries[index].slideNumber}.xml`);
}

test("imageSize reads PNG and JPEG headers, and gives up on anything else", () => {
  assert.deepEqual(imageSize(makePng(640, 360)), { pxWidth: 640, pxHeight: 360 });
  const jpeg = Buffer.from([
    0xff, 0xd8, 0xff, 0xe0, 0x00, 0x04, 0x4a, 0x46, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x02, 0x1c, 0x03, 0xc0, 0x03, 0x01,
    0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01
  ]);
  assert.deepEqual(imageSize(jpeg), { pxWidth: 960, pxHeight: 540 });
  assert.equal(imageSize(Buffer.from("<svg/>")), undefined);
});

test("coverCrop trims the long side evenly, and leaves a matching picture alone", () => {
  assert.deepEqual(coverCrop({ w: 4, h: 2 }, { pxWidth: 400, pxHeight: 100 }), { l: 25000, t: 0, r: 25000, b: 0 });
  assert.deepEqual(coverCrop({ w: 2, h: 4 }, { pxWidth: 100, pxHeight: 100 }), { l: 25000, t: 0, r: 25000, b: 0 });
  assert.deepEqual(coverCrop({ w: 4, h: 1 }, { pxWidth: 100, pxHeight: 100 }), { l: 0, t: 37500, r: 0, b: 37500 });
  assert.equal(coverCrop({ w: 3, h: 2 }, { pxWidth: 1536, pxHeight: 1024 }), undefined);
});

test("a custom slide places the picked variant cropped to fill, or a placeholder when it is missing", async (t) => {
  const dir = await deckProject(t);
  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  const placed: unknown[] = [];
  deck.addCustomSlide(
    new CustomSlide({
      name: "hero",
      draw: async ({ slide, helpers }) => {
        placed.push(await helpers.addImage(slide, { image: "hero" }, { x: 1, y: 1, w: 4, h: 2 }));
        placed.push(await helpers.addImage(slide, { image: "not-yet" }, { x: 5.5, y: 1, w: 3, h: 2 }));
        placed.push(await helpers.addImage(slide, { image: "hero" }, { x: 1, y: 3.5, w: 4, h: 2 }, { fit: "contain" }));
      }
    })
  );

  const report = await deck.render({ output: "output/deck.pptx", report: "output/report.md", progress: false });
  const xml = await slideXml(path.join(dir, "output", "deck.pptx"), 0);

  assert.match(
    xml,
    /<a:srcRect l="25000" r="25000" t="0" b="0"\/>/,
    "a 4:1 picture in a 2:1 box loses a quarter each side"
  );
  assert.match(xml, new RegExp(`<a:ext cx="${emu(4)}" cy="${emu(2)}"/>`), "and fills the box exactly");
  assert.match(xml, /descr="A calm harbour at dawn, boats at rest"/, "the description becomes the alt text");
  assert.match(xml, /Image needed: The team on stage at the summit/);
  assert.deepEqual(placed[2], { x: 1, y: 4, w: 4, h: 1 }, "contain shrinks the box to the picture instead");

  const missing = report.warnings.filter((warning) => warning.code === "image-missing");
  assert.equal(missing.length, 1);
  assert.equal(missing[0].slide, 1);
  assert.match(missing[0].message, /not-yet-1\.jpg\|jpeg\|png/);
  assert.deepEqual(
    report.images.map((record) => [record.id, record.status, record.path]),
    [
      ["hero", "placed", path.join("inputs", "hero-2.png")],
      ["not-yet", "placeholder", undefined],
      ["hero", "placed", path.join("inputs", "hero-2.png")]
    ]
  );

  const markdown = await readFile(path.join(dir, "output", "report.md"), "utf8");
  assert.match(markdown, /## Images\n\n- `hero` \(meant for slide 1\): on slide 1, variant 2 of 2/);
  assert.match(
    markdown,
    /!\[hero 1\]\(\.\.\/inputs\/hero-1\.png\) !\[hero 2 \(picked\)\]\(\.\.\/inputs\/hero-2\.png\)/
  );
  assert.match(markdown, /- `not-yet`: placeholder on slide 1: no file in inputs\/ yet/);
});

test("the addImage override on a cloned slide fills its box, or leaves a captioned placeholder", async (t) => {
  const dir = await deckProject(t);
  await writeFile(path.join(dir, "photo.png"), TINY_PNG);
  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  deck.addSlideFromTemplate({
    templateName: "title-cover",
    variables: { "overline-label": "Cover" },
    overrides: [
      { op: "addImage", id: "hero-art", image: "hero", x: 5, y: 1, w: 4, h: 2 },
      { op: "addImage", id: "team-photo", image: "not-yet", x: 5, y: 3.2, w: 4, h: 2 },
      { op: "addImage", id: "plain", path: "photo.png", x: 1, y: 4, w: 1, h: 1 }
    ]
  });

  const report = await deck.render({ output: "deck.pptx", progress: false });
  const xml = await slideXml(path.join(dir, "deck.pptx"), 0);

  assert.match(xml, /name="hero-art"\/>[\s\S]*?<a:srcRect l="25000" t="0" r="25000" b="0"\/><a:stretch\/>/);
  assert.match(xml, /<p:sp><p:nvSpPr><p:cNvPr id="\d+" name="team-photo"\/>[\s\S]*?Image needed: The team on stage/);
  const plain = xml.match(/<p:pic><p:nvPicPr><p:cNvPr id="\d+" name="plain"\/>[\s\S]*?<\/p:pic>/)?.[0] ?? "";
  assert.ok(plain, "the path-based picture is placed");
  assert.doesNotMatch(plain, /srcRect/, "and, with no fit asked for, stretches exactly as it always has");
  assert.ok(report.warnings.some((warning) => warning.code === "image-missing" && warning.target === "not-yet"));
});

/** A one-slide template holding a real 2x2in picture, for replaceImage to aim at. */
async function templateWithPicture(dir: string): Promise<{ root: string; fieldId: string; name: string }> {
  const seed = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  await writeFile(path.join(dir, "seed.png"), TINY_PNG);
  seed.addSlideFromTemplate({
    templateName: "title-cover",
    variables: { "overline-label": "Seed" },
    overrides: [{ op: "addImage", id: "photo", path: "seed.png", x: 1, y: 1, w: 2, h: 2 }]
  });
  await seed.render({ output: "seed.pptx", progress: false });

  const root = path.join(dir, "library");
  await ingestTemplate({
    source: path.join(dir, "seed.pptx"),
    templateRoot: root,
    templateName: "with-picture",
    slide: 1
  });
  const template = await loadTemplate(root, "with-picture");
  const field = template.fieldsFile.fields.find((candidate) => candidate.type === "image");
  assert.ok(field);
  return { root, fieldId: field.id, name: field.name };
}

test("replaceImage with an image keeps the template's box and crops the picture into it", async (t) => {
  const dir = await deckProject(t);
  const { root, fieldId } = await templateWithPicture(dir);

  const deck = new Presentation({ templateLibrary: root, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  deck.addSlideFromTemplate({
    templateName: "with-picture",
    overrides: [{ op: "replaceImage", target: fieldId, image: "hero" }]
  });
  await deck.render({ output: "replaced.pptx", progress: false });

  const xml = await slideXml(path.join(dir, "replaced.pptx"), 0);
  assert.match(xml, new RegExp(`<a:off x="${emu(1)}" y="${emu(1)}"/><a:ext cx="${emu(2)}" cy="${emu(2)}"/>`));
  assert.match(
    xml,
    /<a:srcRect l="37500" t="0" r="37500" b="0"\/>/,
    "a 4:1 picture in a square keeps its middle quarter"
  );
});

test("replaceImage with a missing image leaves a placeholder that later overrides can still find", async (t) => {
  const dir = await deckProject(t);
  const { root, fieldId, name } = await templateWithPicture(dir);

  const deck = new Presentation({ templateLibrary: root, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  deck.addSlideFromTemplate({
    templateName: "with-picture",
    overrides: [
      { op: "replaceImage", target: fieldId, image: "not-yet" },
      // Aimed at the same field: it must still match after the swap.
      { op: "move", target: fieldId, x: 2, y: 2 }
    ]
  });
  const report = await deck.render({ output: "missing.pptx", progress: false });

  const xml = await slideXml(path.join(dir, "missing.pptx"), 0);
  const placeholder = xml.match(
    new RegExp(`<p:sp><p:nvSpPr><p:cNvPr id="\\d+" name="${name}"/>[\\s\\S]*?</p:sp>`)
  )?.[0];
  assert.ok(placeholder, "the picture became a placeholder with the same name");
  assert.match(placeholder, /Image needed: The team on stage at the summit/);
  assert.match(placeholder, new RegExp(`<a:off x="${emu(2)}" y="${emu(2)}"/>`), "the later move still landed");
  assert.doesNotMatch(xml, new RegExp(`<p:pic>[\\s\\S]*?name="${name}"`), "the template's sample image is gone");
  assert.ok(report.warnings.some((warning) => warning.code === "image-missing"));
});

test("an image the brief does not declare is a build error that suggests the right id", async (t) => {
  const dir = await deckProject(t);
  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  deck.addCustomSlide(
    new CustomSlide({
      name: "typo",
      draw: async ({ slide, helpers }) => {
        await helpers.addImage(slide, { image: "heroo" }, { x: 1, y: 1, w: 2, h: 2 });
      }
    })
  );
  await assert.rejects(
    () => deck.render({ output: "deck.pptx", progress: false }),
    /Image 'heroo' is not declared.*did you mean 'hero'\?/
  );
});

test("a malformed ## Images block warns and draws placeholders instead of failing the build", async (t) => {
  const dir = await deckProject(t);
  await writeFile(path.join(dir, "brief.md"), "## Images\n\n```yaml\n- id: hero\n  varients: 2\n```\n", "utf8");
  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir, assetsDir: BUNDLED_ASSETS });
  deck.addCustomSlide(
    new CustomSlide({
      name: "hero",
      draw: async ({ slide, helpers }) => {
        await helpers.addImage(slide, { image: "hero" }, { x: 1, y: 1, w: 2, h: 2 });
        await helpers.addImage(slide, { image: "hero" }, { x: 4, y: 1, w: 2, h: 2 });
      }
    })
  );

  const report = await deck.render({ output: "deck.pptx", progress: false });
  const invalid = report.warnings.filter((warning) => warning.code === "image-brief-invalid");
  assert.equal(invalid.length, 1, "warned once, however many pictures");
  assert.match(invalid[0].message, /did you mean "variants"/);
  assert.match(await slideXml(path.join(dir, "deck.pptx"), 0), /Image needed: hero/);
});
