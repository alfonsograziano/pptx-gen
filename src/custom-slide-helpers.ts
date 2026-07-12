import { readFile } from "node:fs/promises";
import path from "node:path";
import { C, LAYOUT, LOGO_FILES } from "./brand.js";
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
  rasterizeSvg?: (svg: string) => Promise<Buffer | null>;
}) {
  const { projectDir, assetsDir, shapeType, rasterizeSvg } = options;

  return {
    addHeader(slide: Slide, text: string, opts: { light?: boolean } = {}) {
      const label = text.endsWith("_") ? text.slice(0, -1) : text;
      slide.addText([
        { text: label, options: { color: opts.light === false ? C.white : C.midnight } },
        { text: "_", options: { color: C.green } }
      ], {
        x: LAYOUT.LM,
        y: 0.28,
        w: LAYOUT.CW,
        h: 0.45,
        fontSize: 14,
        fontFace: "Inter",
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
        fontFace: "Inter",
        color: light ? C.mid50 : C.white,
        margin: 0
      });
      this.addNLogo(slide, { light });
    },

    addNLogo(slide: Slide, opts: { light?: boolean } = {}) {
      const file = opts.light ? LOGO_FILES.nMarkColor : LOGO_FILES.nMark;
      slide.addImage({
        path: path.join(assetsDir, file),
        x: 9.3,
        y: 5.0,
        w: 0.33,
        h: 0.26
      });
    },

    addWordmark(slide: Slide, opts: { light?: boolean; x?: number; y?: number; w?: number; h?: number } = {}) {
      const file = opts.light ? LOGO_FILES.wordmarkLight : LOGO_FILES.wordmarkDark;
      slide.addImage({
        path: path.join(assetsDir, file),
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
        fontFace: style.fontFace as string | undefined ?? "Inter",
        color: style.color as string | undefined ?? C.midnight,
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
      const accent = stripHash(opts.accent ?? C.green);
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
        fontFace: "Bitter",
        color: C.midnight,
        margin: 0
      });
      if (opts.body) {
        slide.addText(opts.body, {
          x: opts.x + 0.14,
          y: opts.y + 0.62,
          w: opts.w - 0.28,
          h: Math.max(0.2, opts.h - 0.76),
          fontSize: 9.5,
          fontFace: "Inter",
          color: C.midnight,
          lineSpacingMultiple: LAYOUT.LS,
          valign: "top",
          margin: 0
        });
      }
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
          color: stripHash(opts.color ?? C.mid50),
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
        line: { color: stripHash(opts.color ?? C.midnight), width: strokeWidth, cap: "round" },
        points
      });
    },

    // Embeds a rasterized PNG of the SVG when LibreOffice is available. PptxGenJS
    // embeds SVG images with a "broken image" raster fallback that Google Slides
    // (and older PowerPoint) show instead of the vector, so we ship a real PNG and
    // fall back to the raw SVG only when no rasterizer is available.
    async addSvgDiagram(slide: Slide, opts: Box & { id: string; svg: string }) {
      const png = rasterizeSvg ? await rasterizeSvg(opts.svg) : null;
      const data = png
        ? `image/png;base64,${png.toString("base64")}`
        : `image/svg+xml;base64,${Buffer.from(opts.svg).toString("base64")}`;
      slide.addImage({
        data,
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
        fill: { color: C.midnight },
        line: { color: C.midnight, width: 0 },
        rectRadius: 0.06
      });
      slide.addText(lines.map((line, index) => {
        const trimmed = line.trimStart();
        const isComment = trimmed.startsWith("//") || trimmed.startsWith("#");
        return {
          text: line === "" ? " " : line,
          options: {
            color: isComment ? C.green : C.white,
            breakLine: index < lines.length - 1
          }
        };
      }), {
        x: opts.x + 0.22,
        y: opts.y + 0.14,
        w: opts.w - 0.44,
        h: panelH - 0.28,
        fontSize: 10.5,
        fontFace: opts.fontFace ?? "Consolas",
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
