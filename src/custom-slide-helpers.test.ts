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
