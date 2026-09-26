/**
 * The studio slide: a short message on the left, a generated picture on the
 * right. The text stays native and editable; only the photograph is a picture.
 */
import { CustomSlide, C, FONTS, LAYOUT } from "pptx-gen";

const { LM } = LAYOUT;

// The picture's box: a square against the right margin, above the footer.
const PICTURE = { x: 5.6, y: 0.95, w: 3.9, h: 3.9 };

export function studioSlide(): CustomSlide {
  return new CustomSlide({
    name: "studio-interior",
    background: "light",
    requiredFonts: [FONTS.sans, FONTS.serif],
    async draw({ slide, helpers }) {
      helpers.addHeader(slide, "The studio");

      slide.addText("Room to think, right by the water", {
        x: LM,
        y: 1.05,
        w: 4.4,
        h: 1.1,
        fontSize: 24,
        fontFace: FONTS.serif,
        color: C.ink,
        valign: "top",
        margin: 0
      });

      helpers.addTextBlock(
        slide,
        [
          { text: "One open floor for the whole team.", options: { bullet: true, breakLine: true } },
          { text: "Long tables for working side by side.", options: { bullet: true, breakLine: true } },
          { text: "Daylight all day, from windows on three sides.", options: { bullet: true } }
        ],
        { x: LM, y: 2.35, w: 4.4, h: 1.6 },
        { fontSize: 12, color: C.muted, paraSpaceAfter: 6 }
      );

      // Declared in brief.md's ## Images block. Until `pptx-gen images` has
      // made it, this draws a grey placeholder captioned with the description.
      await helpers.addImage(slide, { image: "studio-interior" }, PICTURE);

      helpers.addFooter(slide);
    }
  });
}
