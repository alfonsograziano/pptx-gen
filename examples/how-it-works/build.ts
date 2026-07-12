/**
 * Build the "How pptx-gen works" explainer slide.
 *
 *   npm run cli -- build --script examples/how-it-works/build.ts
 *
 * The slide is drawn entirely from native shapes, so it is also a live demo of
 * the "design from scratch" mode it explains.
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation } from "../../src/index.js";
import { howItWorksSlide } from "./custom.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));

const deck = new Presentation({
  title: "How pptx-gen works",
  templateLibrary: "templates",
  projectDir: HERE
});

deck.addCustomSlide(howItWorksSlide(1));

await deck.render({
  output: "output/how-it-works.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots"
});
