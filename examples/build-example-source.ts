/**
 * Generate the neutral source deck that the example templates are ingested from.
 *
 * This is a one-off authoring script, not part of the runtime. It draws two
 * plain slides with placeholder text using native shapes, writes them to
 * `examples/example-source.pptx`, and then that file is ingested into
 * `templates/` with the CLI (see the header of each generated template.yml).
 *
 *   npm run cli -- ingest --source examples/example-source.pptx --template title-cover --slide 1
 *   npm run cli -- ingest --source examples/example-source.pptx --template content-lead-bullets --slide 2
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, CustomSlide, C, FONTS, LAYOUT } from "../src/index.js";

const { LM, CW, LS } = LAYOUT;
const HERE = path.dirname(fileURLToPath(import.meta.url));

const titleSlide = new CustomSlide({
  name: "title-cover",
  background: "dark",
  draw({ slide }) {
    slide.addText("Overline label", {
      x: LM, y: 1.35, w: CW, h: 0.35,
      fontSize: 12, fontFace: FONTS.sans, color: C.accent, margin: 0
    });
    slide.addText("Your presentation title goes here", {
      x: LM, y: 1.75, w: CW, h: 1.4,
      fontSize: 34, fontFace: FONTS.sans, color: C.white, margin: 0, valign: "top"
    });
    slide.addText("A short subtitle that sets up the story in one line.", {
      x: LM, y: 3.25, w: CW, h: 0.6,
      fontSize: 14, fontFace: FONTS.sans, color: C.faint, margin: 0
    });
  }
});

const contentSlide = new CustomSlide({
  name: "content-lead-bullets",
  background: "light",
  draw({ slide, helpers }) {
    helpers.addHeader(slide, "Section header");
    slide.addText("A lead statement that frames the three points below.", {
      x: LM, y: 0.95, w: CW, h: 0.9,
      fontSize: 18, fontFace: FONTS.serif, color: C.ink, margin: 0, valign: "top"
    });
    slide.addText([
      { text: "First supporting point that backs up the lead.", options: { bullet: true, breakLine: true } },
      { text: "Second supporting point with a little more detail.", options: { bullet: true, breakLine: true } },
      { text: "Third supporting point to round out the argument.", options: { bullet: true } }
    ], {
      x: LAYOUT.BIL, y: 2.15, w: CW - (LAYOUT.BIL - LM), h: 2.6,
      fontSize: 12, fontFace: FONTS.sans, color: C.ink, lineSpacingMultiple: LS, margin: 0, valign: "top"
    });
    helpers.addFooter(slide, 2, { light: true });
  }
});

const deck = new Presentation({ title: "Example source", projectDir: HERE });
deck.addCustomSlide(titleSlide);
deck.addCustomSlide(contentSlide);

await deck.render({ output: "example-source.pptx", progress: false });
console.log("Wrote examples/example-source.pptx");
