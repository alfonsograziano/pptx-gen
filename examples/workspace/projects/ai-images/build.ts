/**
 * Build the "Our new waterfront studio" demo of generated imagery.
 *
 *   pptx-gen images --project examples/workspace/projects/ai-images   (optional, needs OPENAI_API_KEY)
 *   npm run example
 *
 * The images are declared in brief.md's `## Images` block. This script only
 * places them: it never generates anything, so it builds offline, and a missing
 * image becomes a captioned placeholder rather than an error.
 */
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation } from "pptx-gen";
import { studioSlide } from "./custom.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));

// The title's column: the left half, leaving the right for the picture.
const TEXT_W = 4.7;

const deck = new Presentation({
  title: "Our new waterfront studio",
  projectDir: HERE
});

deck.addSlideFromTemplate({
  templateName: "title-cover",
  variables: {
    "overline-label": "Moving day",
    "your-presentation-title-goes-here": "Our new waterfront studio",
    "a-short-subtitle-that-sets-up-the-st": "More light, more room, and the harbour outside."
  },
  overrides: [
    { op: "resize", target: "overline-label", w: TEXT_W, h: 0.35 },
    { op: "resize", target: "your-presentation-title-goes-here", w: TEXT_W, h: 1.4 },
    { op: "resize", target: "a-short-subtitle-that-sets-up-the-st", w: TEXT_W, h: 0.6 },
    // A full-height picture down the right side, cropped to fill.
    { op: "addImage", id: "harbour-dawn", image: "harbour-dawn", x: 6, y: 0, w: 4, h: 5.625 }
  ]
});

deck.addCustomSlide(studioSlide());

await deck.render({
  output: "output/ai-images.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots"
});
