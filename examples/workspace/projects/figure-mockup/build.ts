/**
 * Build the "figure beside native text" example slide.
 *
 *   npm run cli -- build --script examples/figure-mockup/build.ts
 *
 * The product UI on the right is an HTML figure: `figures/alerts-console.html`,
 * rendered by headless Chrome at build time. Everything else on the slide is
 * native shapes and text. Without a browser installed the figure becomes a
 * captioned placeholder and the deck still builds.
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation } from "pptx-gen";
import { figureMockupSlide } from "./custom.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));

const deck = new Presentation({
  title: "Figures",
  projectDir: HERE
});

deck.addCustomSlide(figureMockupSlide(1));

await deck.render({
  output: "output/figure-mockup.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots"
});
