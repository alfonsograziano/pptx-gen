/**
 * "How pptx-gen works" — a single explainer slide built entirely from native
 * shapes, cards, arrows, and vector icons. It doubles as a live demo of the
 * "design from scratch" mode it describes: open it in PowerPoint or Google
 * Slides and every part is editable and recolorable.
 */
import { CustomSlide, C, FONTS, LAYOUT } from "../../src/index.js";

const { LM, CW } = LAYOUT;

type Tone = "accent" | "accent2" | "accent3" | "ink";

const TONES: Record<Tone, string> = {
  accent: C.accent,
  accent2: C.accent2,
  accent3: C.accent3,
  ink: C.ink
};

// Vertical centre line of the flow, and the tall-card / mode-card geometry.
const CY = 2.95;
const TALL_H = 1.85;
const TALL_Y = CY - TALL_H / 2; // 2.025

const MODE_H = 1.3;
const MODE_TOP_Y = 1.5;
const MODE_BOT_Y = 3.1;
const MODE_TOP_MID = MODE_TOP_Y + MODE_H / 2; // 2.15
const MODE_BOT_MID = MODE_BOT_Y + MODE_H / 2; // 3.75

// Column x/width. Content runs LM (0.75) → 9.5, with gaps for the arrows.
const BRIEF = { x: LM, w: 2.0 };
const MODES = { x: 3.05, w: 2.4 };
const ENGINE = { x: 5.85, w: 1.6 };
const OUTPUT = { x: 7.8, w: 1.7 };

export function howItWorksSlide(pageNum = 1): CustomSlide {
  return new CustomSlide({
    name: "how-pptx-gen-works",
    background: "light",
    requiredFonts: [FONTS.sans, FONTS.serif],
    async draw({ slide, helpers, pptx }) {
      slide.background = { color: C.white };

      helpers.addHeader(slide, "How pptx-gen works");

      slide.addText(
        "Describe a deck to an AI agent. It builds each slide one of two ways, then compiles one editable .pptx.",
        {
          x: LM,
          y: 0.8,
          w: CW,
          h: 0.35,
          fontSize: 11,
          fontFace: FONTS.sans,
          color: C.muted,
          margin: 0
        }
      );

      // Faint container behind the two build modes, to read them as a pair.
      slide.addShape(pptx.ShapeType.roundRect, {
        x: MODES.x - 0.13,
        y: MODE_TOP_Y - 0.16,
        w: MODES.w + 0.26,
        h: MODE_BOT_Y + MODE_H - MODE_TOP_Y + 0.32,
        fill: { color: C.surface },
        line: { color: C.grey30, width: 0.5 },
        rectRadius: 0.08
      });

      // Arrows first, so the card fills tuck the joins under a clean edge.
      const arrow = (from: { x: number; y: number }, to: { x: number; y: number }) =>
        helpers.addArrow(slide, { from, to, color: C.faint, width: 1.4, endArrowType: "triangle" });

      // brief → the two modes (fan out)
      arrow({ x: BRIEF.x + BRIEF.w, y: CY }, { x: MODES.x, y: MODE_TOP_MID });
      arrow({ x: BRIEF.x + BRIEF.w, y: CY }, { x: MODES.x, y: MODE_BOT_MID });
      // the two modes → engine (fan in)
      arrow({ x: MODES.x + MODES.w, y: MODE_TOP_MID }, { x: ENGINE.x, y: CY });
      arrow({ x: MODES.x + MODES.w, y: MODE_BOT_MID }, { x: ENGINE.x, y: CY });
      // engine → output
      arrow({ x: ENGINE.x + ENGINE.w, y: CY }, { x: OUTPUT.x, y: CY });

      // Cards.
      const cards: Array<{
        x: number; y: number; w: number; h: number;
        heading: string; body: string; tone: Tone; icon: string; fill?: string;
      }> = [
        {
          ...BRIEF, y: TALL_Y, h: TALL_H, tone: "accent", icon: "sparkles",
          heading: "Your brief",
          body: "Describe the deck in plain words. An AI agent, driven by the skills, does the work."
        },
        {
          ...MODES, y: MODE_TOP_Y, h: MODE_H, tone: "accent3", icon: "copy",
          heading: "Clone & fill",
          body: "Copy a real slide and fill its text — pixel-identical."
        },
        {
          ...MODES, y: MODE_BOT_Y, h: MODE_H, tone: "accent2", icon: "shapes",
          heading: "Design from scratch",
          body: "Draw native shapes, text and icons — all editable."
        },
        {
          ...ENGINE, y: TALL_Y, h: TALL_H, tone: "ink", icon: "cpu",
          heading: "Engine",
          body: "Assembles slides, embeds fonts, renders previews."
        },
        {
          ...OUTPUT, y: TALL_Y, h: TALL_H, tone: "accent", icon: "presentation",
          heading: "Real .pptx",
          body: "Fully editable. Opens in PowerPoint, Keynote and Google Slides."
        }
      ];

      for (const card of cards) {
        helpers.addCard(slide, {
          x: card.x, y: card.y, w: card.w, h: card.h,
          heading: card.heading,
          body: card.body,
          accent: TONES[card.tone],
          fill: card.fill ?? C.white
        });
      }

      // Icons on top, top-right of each card, tinted to the card's accent.
      for (const card of cards) {
        await helpers.addIcon(
          slide,
          card.icon,
          { x: card.x + card.w - 0.46, y: card.y + 0.15, w: 0.3, h: 0.3 },
          { color: TONES[card.tone] }
        );
      }

      // Grouping caption under the two modes.
      slide.addText("Two modes — mix freely in one deck", {
        x: MODES.x - 0.13,
        y: MODE_BOT_Y + MODE_H + 0.2,
        w: MODES.w + 0.26,
        h: 0.24,
        fontSize: 8.5,
        fontFace: FONTS.sans,
        italic: true,
        color: C.muted,
        align: "center",
        margin: 0
      });

      // Small caption under the output.
      slide.addText("+ screenshots & build report", {
        x: OUTPUT.x,
        y: TALL_Y + TALL_H + 0.14,
        w: OUTPUT.w,
        h: 0.24,
        fontSize: 8.5,
        fontFace: FONTS.sans,
        color: C.faint,
        align: "center",
        margin: 0
      });

      helpers.addFooter(slide, pageNum, { light: true });
    }
  });
}
