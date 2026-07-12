# Custom slide instructions

Use this file when no existing template in `templates/` fits the slide intent.

Prefer cloned templates first. Use custom slides for layouts that need diagrams, flows, arrows, tables, code panels, architecture maps, timelines, or workshop exercises that are not covered by the template library.

## Editable output is the default (read this first)

These decks are opened and edited in **Google Slides** as well as PowerPoint. Google Slides cannot edit or recolor an embedded image (PNG or SVG). It only keeps **native shapes, text, and lines** editable. So:

- **Build everything from native objects.** Use shapes (`addShape`, `addCard`), lines/arrows (`addArrow`, `addConnector`), and text (`addText`, `addTextBlock`). These stay movable and recolorable in Google Slides.
- **Icons must be native, not images.** Use `helpers.addIcon(...)` (icon library) or `helpers.addVectorIcon(slide, svg, box, { color })` (inline SVG). Both convert the vector to a native custom-geometry shape, so the icon can be recolored in Slides by changing its line colour. Do **not** place icons as PNG/SVG images.
- **Do not embed what can be a shape.** `helpers.addSvgDiagram(...)` embeds the SVG as an image (Google Slides shows a broken-image placeholder for embedded SVG, so it is not reliably editable anywhere). Reserve it for genuinely complex vector art (gradients, photos, intricate illustrations) that cannot be expressed as shapes. A box-and-arrow architecture diagram is **not** one of these: build it natively.

When in doubt, ask: "If I open this in Google Slides, can I move and recolor each part?" If the answer is no, rebuild it from native shapes.

## Required workflow

1. Read `design.md`.
2. Inspect the template library and confirm no good template fits.
3. Create `projects/<deck-id>/custom.ts`.
4. Keep all custom layout functions in `custom.ts`.
5. In `build.ts`, call `deck.addCustomSlide(...)`.
6. Run the normal build command and inspect screenshots.

## Design rules

See `design.md` for the full system. In short:

- Slide size is 10 x 5.625 inches.
- Use `LM = 0.75`, `CW = 8.75`, `BIL = 0.95`, `LS = 1.3` (from `LAYOUT`).
- Use the `sans` font for normal text.
- Use the `serif` font only for quotes, callouts, and emphasis.
- Use sentence case.
- Content slides open with a header top-left, in sentence case.
- Use `ink` text on white backgrounds, and white text on `ink` backgrounds.
- Do not put small readable text on a vivid accent fill.
- Use `accent` for accent bars, icons, and rules; use `accent2` / `accent3` for
  diagrams and infographics that need variety.
- Add the page number and logo mark (via `addFooter`) unless the slide is a cover
  or closing slide.

## Project shape

```text
projects/<deck-id>/
  build.ts
  custom.ts
  brief.md
  inputs/
  output/
```

`build.ts` should stay simple:

```ts
import { Presentation } from "../../src/index.js";
import { architectureFlowSlide } from "./custom.js";

const deck = new Presentation({
  title: "AI delivery proposal",
  templateLibrary: "templates",
  projectDir: "projects/<deck-id>",
});

deck.addCustomSlide(architectureFlowSlide({
  pageNum: 3,
  title: "Target architecture_",
}));

await deck.render({
  output: "output/deck.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots",
});
```

`custom.ts` should export named layout functions. Each function returns one `CustomSlide`.

## Example 1: simple callout slide

Use this when the slide needs one strong statement plus short support text.

```ts
import { CustomSlide, C, LAYOUT } from "../../src/index.js";

const { LM, CW, LS } = LAYOUT;

export function calloutSlide(input: {
  pageNum: number;
  title: string;
  callout: string;
  body: string;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-callout",
    requiredFonts: ["Inter", "Lora"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      slide.addText(input.callout, {
        x: LM,
        y: 0.9,
        w: CW,
        h: 1.15,
        fontSize: 18,
        fontFace: "Lora",
        color: C.ink,
        margin: 0,
        valign: "top",
      });

      slide.addText(input.body, {
        x: LM,
        y: 2.2,
        w: CW,
        h: 2.6,
        fontSize: 11,
        fontFace: "Inter",
        color: C.ink,
        lineSpacingMultiple: LS,
        margin: 0,
        valign: "top",
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 2: two-column comparison

Use this for tradeoffs, before and after, current and future state, or two options.

```ts
export function twoColumnSlide(input: {
  pageNum: number;
  title: string;
  left: { heading: string; body: string; bullets?: string[] };
  right: { heading: string; body: string; bullets?: string[] };
}): CustomSlide {
  return new CustomSlide({
    name: "custom-two-column",
    requiredFonts: ["Inter"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const columns = [
        { x: LM, w: 3.9, data: input.left },
        { x: 5.15, w: 4.3, data: input.right },
      ];

      for (const col of columns) {
        helpers.addTextBlock(slide, [
          { text: col.data.heading, options: { bold: true, breakLine: true } },
          { text: col.data.body, options: { breakLine: Boolean(col.data.bullets?.length) } },
          ...(col.data.bullets ?? []).map((item, index, list) => ({
            text: item,
            options: { bullet: true, breakLine: index < list.length - 1 },
          })),
        ], {
          x: col.x,
          y: 0.95,
          w: col.w,
          h: 4.0,
        }, {
          fontSize: 10.5,
          fontFace: "Inter",
          color: C.ink,
          lineSpacingMultiple: LS,
        });
      }

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 3: three card slide

Use this for pillars, workstreams, principles, or three related steps.

```ts
export function threeCardSlide(input: {
  pageNum: number;
  title: string;
  intro: string;
  cards: Array<{ heading: string; body: string }>;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-three-card",
    requiredFonts: ["Inter", "Lora"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      slide.addText(input.intro, {
        x: LM,
        y: 0.88,
        w: CW,
        h: 0.75,
        fontSize: 16,
        fontFace: "Lora",
        color: C.ink,
        margin: 0,
      });

      const cardW = 2.63;
      const gutter = 0.22;

      input.cards.slice(0, 3).forEach((card, index) => {
        const x = LM + index * (cardW + gutter);
        helpers.addCard(slide, {
          x,
          y: 1.95,
          w: cardW,
          h: 2.95,
          heading: card.heading,
          body: card.body,
          accent: C.accent,
        });
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 4: table slide

Use this for structured data. Keep tables small enough to read.

```ts
export function tableSlide(input: {
  pageNum: number;
  title: string;
  rows: string[][];
  colW?: number[];
}): CustomSlide {
  return new CustomSlide({
    name: "custom-table",
    requiredFonts: ["Inter", "Lora"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const styledRows = input.rows.map((row, rowIndex) =>
        row.map((cell) => rowIndex === 0
          ? {
              text: cell,
              options: {
                bold: true,
                fontFace: "Lora",
                fontSize: 12,
                color: C.ink,
                fill: { color: C.accent },
                valign: "middle",
              },
            }
          : { text: cell, options: { valign: "middle" } }
        )
      );

      slide.addTable(styledRows, {
        x: LM,
        y: 0.95,
        w: CW,
        colW: input.colW,
        border: { pt: 0.5, color: C.grey30 },
        fontSize: 9.5,
        fontFace: "Inter",
        color: C.ink,
        fill: { color: C.white },
        rowH: 0.4,
        margin: [3, 6, 3, 6],
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 5: icon grid

Use this for takeaways, capabilities, or controls.

```ts
export function iconGridSlide(input: {
  pageNum: number;
  title: string;
  items: Array<{ icon: string; heading: string; body: string }>;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-icon-grid",
    requiredFonts: ["Inter", "Lora"],
    async draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const colW = 1.5;
      const gutter = 0.3125;

      for (const [index, item] of input.items.slice(0, 5).entries()) {
        const x = LM + index * (colW + gutter);

        await helpers.addIcon(slide, item.icon, {
          x: x + 0.475,
          y: 1.0,
          w: 0.55,
          h: 0.55,
        }, {
          color: C.ink,
        });

        slide.addText(item.heading, {
          x,
          y: 1.68,
          w: colW,
          h: 0.7,
          fontSize: 13,
          fontFace: "Lora",
          color: C.ink,
          margin: 0,
        });

        slide.addText(item.body, {
          x,
          y: 2.5,
          w: colW,
          h: 2.1,
          fontSize: 9,
          fontFace: "Inter",
          color: C.ink,
          lineSpacingMultiple: LS,
          margin: 0,
        });

        if (index < input.items.length - 1) {
          slide.addShape("line", {
            x: x + colW + gutter / 2,
            y: 1.0,
            w: 0,
            h: 3.6,
            line: { color: C.grey30, width: 1, dashType: "dash" },
          });
        }
      }

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 6: timeline or plan

Use this for roadmap, delivery plan, or sprint sequence slides.

```ts
export function timelineSlide(input: {
  pageNum: number;
  title: string;
  columns: string[];
  phases: Array<{
    label: string;
    start: number;
    span: number;
    dark?: boolean;
  }>;
  milestones: Array<{
    label: string;
    colOffset: number;
  }>;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-timeline",
    requiredFonts: ["Inter"],
    draw({ slide, helpers, pptx }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const tableX = LM;
      const tableY = 0.95;
      const colW = CW / input.columns.length;
      const headerH = 0.42;

      input.columns.forEach((label, index) => {
        slide.addShape(pptx.ShapeType.rect, {
          x: tableX + index * colW,
          y: tableY,
          w: colW,
          h: headerH,
          fill: { color: C.accent },
          line: { color: C.accent, width: 0 },
        });
        slide.addText(label, {
          x: tableX + index * colW + 0.08,
          y: tableY + 0.06,
          w: colW - 0.1,
          h: headerH - 0.1,
          fontSize: 10,
          fontFace: "Inter",
          bold: true,
          color: C.ink,
          margin: 0,
        });
      });

      for (let i = 1; i < input.columns.length; i++) {
        slide.addShape(pptx.ShapeType.line, {
          x: tableX + i * colW,
          y: tableY + headerH,
          w: 0,
          h: 3.05,
          line: { color: C.grey30, width: 1, dashType: "dash" },
        });
      }

      const rowH = 0.66;
      const rowStartY = tableY + headerH + 0.15;

      input.phases.forEach((phase, index) => {
        const y = rowStartY + index * rowH;
        const x = tableX + phase.start * colW + 0.08;
        const w = phase.span * colW - 0.16;
        const fill = phase.dark ? C.ink : C.accent;

        slide.addShape(pptx.ShapeType.roundRect, {
          x,
          y: y + 0.12,
          w,
          h: 0.4,
          fill: { color: fill },
          line: { color: fill, width: 0 },
          radius: 0.06,
        });
        slide.addText(phase.label, {
          x: x + 0.12,
          y: y + 0.12,
          w: w - 0.2,
          h: 0.4,
          fontSize: 9.5,
          fontFace: "Inter",
          bold: true,
          color: phase.dark ? C.white : C.ink,
          valign: "mid",
          margin: 0,
        });
      });

      const dotY = rowStartY + input.phases.length * rowH + 0.05;
      slide.addShape(pptx.ShapeType.line, {
        x: tableX,
        y: dotY + 0.18,
        w: CW,
        h: 0,
        line: { color: C.faint, width: 1 },
      });

      input.milestones.forEach((milestone) => {
        const x = tableX + milestone.colOffset * colW + 0.1;
        slide.addShape(pptx.ShapeType.ellipse, {
          x: x - 0.1,
          y: dotY + 0.03,
          w: 0.3,
          h: 0.3,
          fill: { color: C.white },
          line: { color: C.ink, width: 1.5 },
        });
        slide.addText(milestone.label, {
          x: x - 0.55,
          y: dotY + 0.38,
          w: 1.3,
          h: 0.3,
          fontSize: 8,
          fontFace: "Inter",
          color: C.muted,
          align: "center",
          margin: 0,
        });
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 7: architecture diagram with arrows

Use this for systems, agent flows, data pipelines, or process maps.

```ts
export function architectureFlowSlide(input: {
  pageNum: number;
  title: string;
  nodes: Array<{
    id: string;
    label: string;
    detail?: string;
    x: number;
    y: number;
    w?: number;
    h?: number;
    tone?: "accent" | "accent2" | "accent3" | "ink";
  }>;
  arrows: Array<{
    from: string;
    to: string;
    label?: string;
    dashed?: boolean;
  }>;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-architecture-flow",
    requiredFonts: ["Inter"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const nodesById = new Map(input.nodes.map((node) => [
        node.id,
        { ...node, w: node.w ?? 1.7, h: node.h ?? 0.72 },
      ]));

      for (const arrow of input.arrows) {
        const from = nodesById.get(arrow.from);
        const to = nodesById.get(arrow.to);
        if (!from || !to) continue;

        helpers.addArrow(slide, {
          from: {
            x: from.x + from.w,
            y: from.y + from.h / 2,
          },
          to: {
            x: to.x,
            y: to.y + to.h / 2,
          },
          color: C.muted,
          width: 1.2,
          dashed: arrow.dashed,
          endArrowType: "triangle",
        });

        if (arrow.label) {
          slide.addText(arrow.label, {
            x: (from.x + to.x) / 2 - 0.45,
            y: from.y - 0.25,
            w: 0.9,
            h: 0.22,
            fontSize: 7.5,
            fontFace: "Inter",
            color: C.muted,
            align: "center",
            margin: 0,
          });
        }
      }

      for (const node of nodesById.values()) {
        const accent = node.tone === "accent3"
          ? C.accent3
          : node.tone === "accent2"
            ? C.accent2
            : node.tone === "ink"
              ? C.ink
              : C.accent;

        helpers.addCard(slide, {
          x: node.x,
          y: node.y,
          w: node.w,
          h: node.h,
          heading: node.label,
          body: node.detail,
          accent,
          fill: C.surface,
        });
      }

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 8: SVG diagram fallback (rasterized, NOT editable in Google Slides)

> **Last resort.** `addSvgDiagram` rasterizes the SVG to a PNG, so the result is a flat image that cannot be moved or recolored in Google Slides. Use it only for complex vector art that cannot be built from native shapes. For boxes, arrows, flows, and icons, use native shapes (Example 7) and `addVectorIcon`/`addIcon` instead.

```ts
export function svgLoopSlide(input: {
  pageNum: number;
  title: string;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-svg-loop",
    requiredFonts: ["Inter"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      const svg = `
        <svg xmlns="http://www.w3.org/2000/svg" width="900" height="330" viewBox="0 0 900 330">
          <defs>
            <marker id="arrow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto">
              <path d="M0,0 L0,6 L9,3 z" fill="#526288"/>
            </marker>
          </defs>
          <rect x="20" y="90" width="190" height="80" rx="8" fill="#F4F8FA" stroke="#00E5A4" stroke-width="5"/>
          <rect x="355" y="90" width="190" height="80" rx="8" fill="#F4F8FA" stroke="#166BFF" stroke-width="5"/>
          <rect x="690" y="90" width="190" height="80" rx="8" fill="#F4F8FA" stroke="#8950FF" stroke-width="5"/>
          <text x="115" y="137" text-anchor="middle" font-family="Inter" font-size="24" fill="#000E38">Prompt</text>
          <text x="450" y="137" text-anchor="middle" font-family="Inter" font-size="24" fill="#000E38">Evaluate</text>
          <text x="785" y="137" text-anchor="middle" font-family="Inter" font-size="24" fill="#000E38">Improve</text>
          <path d="M215 130 C270 130 300 130 350 130" fill="none" stroke="#526288" stroke-width="3" marker-end="url(#arrow)"/>
          <path d="M550 130 C605 130 635 130 685 130" fill="none" stroke="#526288" stroke-width="3" marker-end="url(#arrow)"/>
          <path d="M785 180 C700 280 190 280 115 180" fill="none" stroke="#526288" stroke-width="3" stroke-dasharray="8 8" marker-end="url(#arrow)"/>
        </svg>
      `;

      helpers.addSvgDiagram(slide, {
        id: "eval-loop",
        svg,
        x: 0.75,
        y: 1.25,
        w: 8.75,
        h: 3.2,
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 9: code panel

Use this for API snippets, scorer examples, JSON contracts, or configuration.

```ts
export function codePanelSlide(input: {
  pageNum: number;
  title: string;
  intro?: string;
  code: string;
  caption?: string;
}): CustomSlide {
  return new CustomSlide({
    name: "custom-code-panel",
    requiredFonts: ["Inter"],
    draw({ slide, helpers }) {
      slide.background = { color: C.white };
      helpers.addHeader(slide, input.title);

      let panelY = 0.95;
      if (input.intro) {
        slide.addText(input.intro, {
          x: LM,
          y: 0.86,
          w: CW,
          h: 0.55,
          fontSize: 11.5,
          fontFace: "Inter",
          color: C.ink,
          lineSpacingMultiple: LS,
          margin: 0,
        });
        panelY = 1.5;
      }

      helpers.addCodePanel(slide, {
        code: input.code,
        x: LM,
        y: panelY,
        w: CW,
        maxH: 5.12 - panelY,
        fontFace: "JetBrains Mono",
      });

      if (input.caption) {
        slide.addText(input.caption, {
          x: LM,
          y: 5.0,
          w: CW,
          h: 0.3,
          fontSize: 8,
          fontFace: "Inter",
          italic: true,
          color: C.muted,
          margin: 0,
        });
      }

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Example 10: exercise slide

Use this for workshops and interactive sessions.

```ts
export function exerciseSlide(input: {
  pageNum: number;
  title: string;
  tag: string;
  time?: string;
  task: string;
  steps: string[];
}): CustomSlide {
  return new CustomSlide({
    name: "custom-exercise",
    requiredFonts: ["Inter", "Lora"],
    draw({ slide, helpers, pptx }) {
      slide.background = { color: C.white };

      slide.addShape(pptx.ShapeType.rect, {
        x: 0,
        y: 0,
        w: 0.18,
        h: 5.625,
        fill: { color: C.accent },
        line: { color: C.accent, width: 0 },
      });

      helpers.addHeader(slide, input.title);

      const tagText = input.tag + (input.time ? `   |   ${input.time}` : "");
      slide.addShape(pptx.ShapeType.roundRect, {
        x: LM,
        y: 0.92,
        w: 3.0,
        h: 0.36,
        fill: { color: C.accent },
        line: { color: C.accent, width: 0 },
        radius: 0.09,
      });
      slide.addText(tagText, {
        x: LM,
        y: 0.92,
        w: 3.0,
        h: 0.36,
        fontSize: 9,
        fontFace: "Inter",
        bold: true,
        color: C.ink,
        align: "center",
        valign: "mid",
        margin: 0,
      });

      slide.addText(input.task, {
        x: LM,
        y: 1.5,
        w: CW,
        h: 1.0,
        fontSize: 19,
        fontFace: "Lora",
        color: C.ink,
        margin: 0,
      });

      slide.addText(input.steps.map((step, index) => ({
        text: step,
        options: {
          bullet: true,
          breakLine: index < input.steps.length - 1,
        },
      })), {
        x: BIL,
        y: 2.7,
        w: CW - (BIL - LM),
        h: 2.35,
        fontSize: 11,
        fontFace: "Inter",
        color: C.ink,
        lineSpacingMultiple: LS,
        margin: 0,
      });

      helpers.addFooter(slide, input.pageNum, { light: true });
    },
  });
}
```

## Diagram guidance

- Build diagrams from native shapes, lines, and text so they stay editable in Google Slides. Do not render a diagram as one SVG/PNG image.
- For real-world imagery the tool cannot draw (a photo, screenshot, logo, or chart), do not fake it: drop `helpers.addImagePlaceholder(slide, { x, y, w, h, caption })` — a grey box with a centered italic caption describing what belongs there — so the user can supply the asset later. Use it sparingly and write a specific caption; build anything that can be a shape natively instead.
- Use `addVectorIcon` / `addIcon` for icons (native custom geometry), never icons-as-images.
- Use boxes for systems, services, actors, or steps.
- Use arrows for data flow, control flow, or sequence.
- Label arrows only when the meaning is not obvious.
- Prefer left-to-right flow for architecture diagrams.
- Use dashed arrows for optional, async, feedback, or monitoring paths.
- Use `accent2` and `accent3` to separate secondary lanes or domains.
- Keep nodes large enough to edit in PowerPoint.
- Do not draw dense diagrams that need zooming.

## Quality checklist

- The slide has a clear job.
- The title is sentence case.
- Text is legible at presentation size.
- Nothing overlaps the footer.
- Arrows touch the right visual targets.
- **Every part is a native shape, line, text, or vector icon, so it is editable and recolorable in Google Slides.** No diagrams or icons embedded as images, unless it is complex art that cannot be a shape (then it is a deliberate, noted exception).
- Icons use `addIcon` / `addVectorIcon` (native), not image placement.
- No random values, current dates, or network calls are used.
- Screenshots are inspected after build.
