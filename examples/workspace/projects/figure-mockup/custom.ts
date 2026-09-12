import { CustomSlide, type Figure } from "pptx-gen";

// A figure beside native text. The product UI on the right is an HTML figure,
// because a mocked-up app screen has no parts a reader would want to click and
// recolor. Everything else — the header, the points, the footer — stays native,
// so it is editable, searchable, and rendered with the deck's embedded fonts.

const LM = 0.75;
const FIGURE_BOX = { x: 5.15, y: 1.62, w: 4.3, h: 2.408 };

const CONSOLE_MOCKUP: Figure = {
  id: "alerts-console",
  htmlFile: "figures/alerts-console.html",
  caption: "The alerts console, with the checkout latency alert firing",
  // 1000 x 560 is the same 1.79:1 shape as the box above, so the figure fills it
  // exactly instead of being shrunk to fit.
  viewport: { width: 1000, height: 560 }
};

export function figureMockupSlide(pageNum = 1): CustomSlide {
  return new CustomSlide({
    name: "figure-mockup",
    draw: async ({ slide, helpers, design }) => {
      const { colors } = design;

      helpers.addHeader(slide, "What the operator sees");

      helpers.addTextBlock(
        slide,
        [{ text: "One screen, three states" }],
        { x: LM, y: 1.2, w: 4.1, h: 0.5 },
        { fontSize: 20, color: colors.ink }
      );

      helpers.addTextBlock(
        slide,
        [
          {
            text: "Every incident lands in one queue, so the on-call engineer never has to decide where to look first."
          }
        ],
        { x: LM, y: 1.95, w: 4.0, h: 0.8 },
        { fontSize: 11, color: colors.muted }
      );

      const points = [
        "Firing alerts sort to the top, with the owning team attached",
        "Acknowledging an alert keeps it visible instead of hiding it",
        "Resolved alerts stay for the retro, then age out after a week"
      ];
      points.forEach((point, index) => {
        slide.addShape("rect", {
          x: LM,
          y: 2.95 + index * 0.55 + 0.08,
          w: 0.1,
          h: 0.1,
          fill: { color: colors.accent },
          line: { color: colors.accent, width: 0 }
        });
        helpers.addTextBlock(
          slide,
          [{ text: point }],
          { x: LM + 0.26, y: 2.95 + index * 0.55, w: 3.84, h: 0.5 },
          { fontSize: 10, color: colors.ink }
        );
      });

      // The figure. If no browser is installed this draws the captioned
      // placeholder instead and the slide still makes its point.
      await helpers.addFigure(slide, CONSOLE_MOCKUP, FIGURE_BOX);

      // A caption that says something the figure does not, rather than repeating
      // its caption — that one is the stand-in text if the figure cannot render.
      helpers.addTextBlock(
        slide,
        [{ text: "Incident queue \u00b7 3 open, 1 firing" }],
        { x: FIGURE_BOX.x, y: FIGURE_BOX.y + FIGURE_BOX.h + 0.12, w: FIGURE_BOX.w, h: 0.4 },
        { fontSize: 9, color: colors.muted, italic: true }
      );

      helpers.addFooter(slide, pageNum);
    }
  });
}
