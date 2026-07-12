import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { C, FONTS, LAYOUT, LOGO_FILES } from "./design.js";
import { parseSvg, svgToGeomPoints } from "./svg-path.js";

type Slide = {
  addText: (text: string | TextRun[], options?: Record<string, unknown>) => unknown;
  addImage: (options: Record<string, unknown>) => unknown;
  addShape: (shapeName: string, options?: Record<string, unknown>) => unknown;
};

type ShapeType = {
  rect: string;
  line: string;
  roundRect: string;
};

type TextRun = {
  text: string;
  options?: Record<string, unknown>;
};

type TextOptions = Record<string, unknown>;
type ArrowType = "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";

export type Box = {
  x: number;
  y: number;
  w: number;
  h: number;
};

export type Point = {
  x: number;
  y: number;
};

export type CustomSlideHelpers = ReturnType<typeof createCustomSlideHelpers>;

export function createCustomSlideHelpers(options: {
  projectDir: string;
  assetsDir: string;
  shapeType: ShapeType;
}) {
  const { projectDir, assetsDir, shapeType } = options;

  return {
    addHeader(slide: Slide, text: string, opts: { light?: boolean } = {}) {
      slide.addText(text, {
        x: LAYOUT.LM,
        y: 0.28,
        w: LAYOUT.CW,
        h: 0.45,
        fontSize: 14,
        fontFace: FONTS.sans,
        color: opts.light === false ? C.white : C.ink,
        margin: 0
      });
    },

    addFooter(slide: Slide, pageNum: number, opts: { light?: boolean } = {}) {
      const light = opts.light ?? true;
      slide.addText([
        { text: "-", options: { breakLine: true } },
        { text: String(pageNum) }
      ], {
        x: 0.5,
        y: 5.1,
        w: 0.35,
        h: 0.38,
        fontSize: 7,
        fontFace: FONTS.sans,
        color: light ? C.muted : C.white,
        margin: 0
      });
      this.addLogo(slide, { light });
    },

    // Places the small logo mark bottom-right, if the configured PNG exists in
    // `assets/`. Ships as a no-op until you add your own logo files.
    addLogo(slide: Slide, opts: { light?: boolean } = {}) {
      const file = opts.light ? LOGO_FILES.markDark : LOGO_FILES.markLight;
      const logoPath = path.join(assetsDir, file);
      if (!existsSync(logoPath)) return;
      slide.addImage({ path: logoPath, x: 9.3, y: 5.0, w: 0.33, h: 0.26 });
    },

    // Places the full wordmark, if the configured PNG exists in `assets/`.
    addWordmark(slide: Slide, opts: { light?: boolean; x?: number; y?: number; w?: number; h?: number } = {}) {
      const file = opts.light ? LOGO_FILES.wordmarkDark : LOGO_FILES.wordmarkLight;
      const logoPath = path.join(assetsDir, file);
      if (!existsSync(logoPath)) return;
      slide.addImage({
        path: logoPath,
        x: opts.x ?? LAYOUT.LM,
        y: opts.y ?? 0.62,
        w: opts.w ?? 1.8,
        h: opts.h ?? 0.31
      });
    },

    addTextBlock(
      slide: Slide,
      runs: TextRun[],
      box: Box,
      style: TextOptions = {}
    ) {
      slide.addText(runs, {
        ...box,
        fontSize: style.fontSize as number | undefined ?? 10,
        fontFace: style.fontFace as string | undefined ?? FONTS.sans,
        color: style.color as string | undefined ?? C.ink,
        lineSpacingMultiple: style.lineSpacingMultiple as number | undefined ?? LAYOUT.LS,
        valign: style.valign as string | undefined ?? "top",
        margin: style.margin as number | undefined ?? 0,
        ...style
      });
    },

    addCard(slide: Slide, opts: Box & {
      heading: string;
      body?: string;
      accent?: string;
      fill?: string;
    }) {
      const accent = stripHash(opts.accent ?? C.accent);
      const fill = stripHash(opts.fill ?? C.white);
      slide.addShape(shapeType.rect, {
        x: opts.x,
        y: opts.y,
        w: opts.w,
        h: opts.h,
        fill: { color: fill },
        line: { color: C.grey30, width: 0.5 }
      });
      slide.addShape(shapeType.line, {
        x: opts.x,
        y: opts.y,
        w: opts.w * 0.45,
        h: 0,
        line: { color: accent, width: 4 }
      });
      slide.addText(opts.heading, {
        x: opts.x + 0.14,
        y: opts.y + 0.18,
        w: opts.w - 0.28,
        h: 0.34,
        fontSize: 12,
        fontFace: FONTS.serif,
        color: C.ink,
        margin: 0
      });
      if (opts.body) {
        slide.addText(opts.body, {
          x: opts.x + 0.14,
          y: opts.y + 0.62,
          w: opts.w - 0.28,
          h: Math.max(0.2, opts.h - 0.76),
          fontSize: 9.5,
          fontFace: FONTS.sans,
          color: C.ink,
          lineSpacingMultiple: LAYOUT.LS,
          valign: "top",
          margin: 0
        });
      }
    },

    // A native stand-in for a photo or graphic the user should supply later.
    // Draws a soft grey box with a dashed border and a centered italic caption
    // describing what belongs there. It stays editable in PowerPoint and Google
    // Slides. Use it only when a real image genuinely helps and cannot be built
    // from native shapes; do not overuse it.
    addImagePlaceholder(slide: Slide, opts: Box & { caption: string }) {
      slide.addShape(shapeType.roundRect, {
        x: opts.x,
        y: opts.y,
        w: opts.w,
        h: opts.h,
        fill: { color: C.grey10 },
        line: { color: C.grey30, width: 1, dashType: "dash" },
        rectRadius: 0.06
      });
      slide.addText(opts.caption, {
        x: opts.x + 0.2,
        y: opts.y,
        w: opts.w - 0.4,
        h: opts.h,
        fontSize: 10,
        fontFace: FONTS.sans,
        italic: true,
        color: C.muted,
        align: "center",
        valign: "middle",
        margin: 0
      });
    },

    addArrow(slide: Slide, opts: {
      from: Point;
      to: Point;
      color?: string;
      width?: number;
      dashed?: boolean;
      beginArrowType?: ArrowType;
      endArrowType?: ArrowType;
    }) {
      slide.addShape(shapeType.line, {
        x: opts.from.x,
        y: opts.from.y,
        w: opts.to.x - opts.from.x,
        h: opts.to.y - opts.from.y,
        line: {
          color: stripHash(opts.color ?? C.muted),
          width: opts.width ?? 1.2,
          dashType: opts.dashed ? "dash" : "solid",
          beginArrowType: opts.beginArrowType ?? "none",
          endArrowType: opts.endArrowType ?? "triangle"
        }
      });
    },

    addConnector(slide: Slide, opts: {
      points: Point[];
      color?: string;
      width?: number;
      dashed?: boolean;
      endArrowType?: ArrowType;
    }) {
      for (let index = 1; index < opts.points.length; index += 1) {
        this.addArrow(slide, {
          from: opts.points[index - 1],
          to: opts.points[index],
          color: opts.color,
          width: opts.width,
          dashed: opts.dashed,
          endArrowType: index === opts.points.length - 1 ? opts.endArrowType ?? "triangle" : "none"
        });
      }
    },

    async addIcon(slide: Slide, icon: string, box: Box, opts: { color?: string; width?: number } = {}) {
      const iconPath = resolveAsset(projectDir, assetsDir, icon, "icons");
      const svg = await readFile(iconPath, "utf8");
      this.addVectorIcon(slide, svg, box, opts);
    },

    // Render an SVG (stroke-based icon or simple vector art) as a NATIVE custom
    // geometry shape. This stays editable and recolorable in PowerPoint AND
    // Google Slides (change the line colour), unlike a rasterized/SVG image.
    // Use this for icons; reserve addSvgDiagram for complex art that cannot be
    // expressed as shapes.
    addVectorIcon(slide: Slide, svg: string, box: Box, opts: { color?: string; width?: number } = {}) {
      const parsed = parseSvg(svg);
      const points = svgToGeomPoints(parsed, box.w, box.h);
      const strokeWidth = opts.width ?? Math.max(0.5, (2 / parsed.vbH) * box.h * 72);
      // Omit `fill` entirely so PptxGenJS emits <a:noFill/>; passing {type:"none"}
      // is truthy and leaves the shape with a default (theme) fill instead.
      slide.addShape("custGeom", {
        x: box.x,
        y: box.y,
        w: box.w,
        h: box.h,
        line: { color: stripHash(opts.color ?? C.ink), width: strokeWidth, cap: "round" },
        points
      });
    },

    // Embeds the SVG as an image. PptxGenJS embeds SVG with a raster fallback
    // that Google Slides and older PowerPoint show as a broken-image placeholder,
    // so the result is not reliably editable and may not render everywhere.
    // Strongly prefer native shapes plus addVectorIcon/addIcon; reserve this for
    // complex art (gradients, illustrations) that cannot be expressed as shapes.
    addSvgDiagram(slide: Slide, opts: Box & { id: string; svg: string }) {
      slide.addImage({
        data: `image/svg+xml;base64,${Buffer.from(opts.svg).toString("base64")}`,
        x: opts.x,
        y: opts.y,
        w: opts.w,
        h: opts.h
      });
    },

    addCodePanel(slide: Slide, opts: {
      code: string;
      x: number;
      y: number;
      w: number;
      maxH: number;
      fontFace?: string;
    }) {
      const lines = opts.code.split("\n");
      const panelH = Math.min(opts.maxH, lines.length * 0.225 + 0.34);
      slide.addShape(shapeType.roundRect, {
        x: opts.x,
        y: opts.y,
        w: opts.w,
        h: panelH,
        fill: { color: C.ink },
        line: { color: C.ink, width: 0 },
        rectRadius: 0.06
      });
      slide.addText(lines.map((line, index) => {
        const trimmed = line.trimStart();
        const isComment = trimmed.startsWith("//") || trimmed.startsWith("#");
        return {
          text: line === "" ? " " : line,
          options: {
            color: isComment ? C.accent : C.white,
            breakLine: index < lines.length - 1
          }
        };
      }), {
        x: opts.x + 0.22,
        y: opts.y + 0.14,
        w: opts.w - 0.44,
        h: panelH - 0.28,
        fontSize: 10.5,
        fontFace: opts.fontFace ?? FONTS.mono,
        lineSpacingMultiple: 1.18,
        valign: "top",
        margin: 0
      });
    }
  };
}

function resolveAsset(projectDir: string, assetsDir: string, filePath: string, fallbackFolder?: string): string {
  if (path.isAbsolute(filePath)) return filePath;
  if (filePath.includes("/") || filePath.includes("\\")) return path.resolve(projectDir, filePath);
  return path.join(assetsDir, fallbackFolder ?? "", filePath.endsWith(".svg") ? filePath : `${filePath}.svg`);
}

function stripHash(value: string): string {
  return value.replace(/^#/, "");
}
