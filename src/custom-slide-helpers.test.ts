import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, rm } from "node:fs/promises";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";
import { C } from "./design.js";
import { STARTER_TEMPLATES } from "./test-fixtures.js";

/** How many runs or shapes on the slide carry this colour. */
function countColor(xml: string, color: string): number {
  return xml.split(`<a:srgbClr val="${color.replace(/^#/, "").toUpperCase()}"/>`).length - 1;
}

// One render covers both cards: building a deck is the slow part of this file,
// and the two cards cannot collide because every colour below is distinct.
test("addCard keeps ink text by default and takes a colour for dark fills", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-helpers-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir });
  deck.addCustomSlide(
    new CustomSlide({
      name: "cards",
      background: { color: "334455" },
      draw({ slide, helpers }) {
        // Default: unchanged from before the option existed.
        helpers.addCard(slide, {
          x: 0.5,
          y: 1,
          w: 3,
          h: 1.5,
          heading: "Ingest",
          body: "Clone real slides.",
          fill: C.surface,
          accent: C.accent
        });
        // A raised panel on a dark slide: dark fill, light text.
        helpers.addCard(slide, {
          x: 4,
          y: 1,
          w: 3,
          h: 1.5,
          heading: "Compile",
          body: "Fill fields and merge.",
          fill: C.inkSoft,
          accent: C.accentOnDark,
          color: C.white
        });
      }
    })
  );
  await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const entries = await getSlideEntries(pkg);
  const xml = await pkg.text(`ppt/slides/slide${entries[0].slideNumber}.xml`);

  // Heading and body of the default card, and nothing else on the slide.
  assert.equal(countColor(xml, C.ink), 2);
  // Heading and body of the dark card. Before `color` existed these were ink
  // too, so a card on a dark slide was unreadable and had to be hand-rolled.
  assert.equal(countColor(xml, C.white), 2);
});

// pptxgenjs writes a negative width or height through instead of flipping the
// shape, so before the fix an arrow pointing left or up had a negative extent:
// invalid in PowerPoint, and drawn the wrong way round where it rendered at all.
test("addArrow points the way it is drawn in every direction", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-helpers-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const from = { x: 3, y: 2 };
  const cases = [
    { name: "right", to: { x: 4, y: 2 }, box: { x: 3, y: 2, w: 1, h: 0 }, flipH: false, flipV: false },
    { name: "left", to: { x: 2, y: 2 }, box: { x: 2, y: 2, w: 1, h: 0 }, flipH: true, flipV: false },
    { name: "down", to: { x: 3, y: 3 }, box: { x: 3, y: 2, w: 0, h: 1 }, flipH: false, flipV: false },
    { name: "up", to: { x: 3, y: 1 }, box: { x: 3, y: 1, w: 0, h: 1 }, flipH: false, flipV: true },
    { name: "up-left", to: { x: 2, y: 1 }, box: { x: 2, y: 1, w: 1, h: 1 }, flipH: true, flipV: true }
  ];

  const deck = new Presentation({ templateLibrary: STARTER_TEMPLATES, projectDir: dir });
  deck.addCustomSlide(
    new CustomSlide({
      name: "arrows",
      draw({ slide, helpers }) {
        for (const { to } of cases) helpers.addArrow(slide, { from, to });
        // A connector that doubles back runs through the same code for each leg.
        helpers.addConnector(slide, {
          points: [
            { x: 6, y: 1 },
            { x: 8, y: 1 },
            { x: 8, y: 3 },
            { x: 6, y: 3 }
          ]
        });
      }
    })
  );
  await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const entries = await getSlideEntries(pkg);
  const xml = await pkg.text(`ppt/slides/slide${entries[0].slideNumber}.xml`);

  const EMU = 914400;
  const lines = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)]
    .map((match) => match[0])
    .filter((shape) => shape.includes('prst="line"'))
    .map((shape) => {
      const xfrm = shape.match(/<a:xfrm([^>]*)><a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(-?\d+)" cy="(-?\d+)"\/>/);
      assert.ok(xfrm, "every line has a transform");
      const [, attrs, x, y, cx, cy] = xfrm;
      return {
        box: { x: Number(x) / EMU, y: Number(y) / EMU, w: Number(cx) / EMU, h: Number(cy) / EMU },
        flipH: attrs.includes('flipH="1"'),
        flipV: attrs.includes('flipV="1"')
      };
    });

  assert.equal(lines.length, cases.length + 3);
  for (const [index, expected] of cases.entries()) {
    assert.deepEqual(
      lines[index],
      { box: expected.box, flipH: expected.flipH, flipV: expected.flipV },
      `arrow pointing ${expected.name}`
    );
  }
  // The connector's third leg runs right to left.
  assert.deepEqual(lines[cases.length + 2], { box: { x: 6, y: 3, w: 2, h: 0 }, flipH: true, flipV: false });
  for (const line of lines) assert.ok(line.box.w >= 0 && line.box.h >= 0, "no line has a negative extent");
});
