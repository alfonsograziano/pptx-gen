import { readFile } from "node:fs/promises";
import { C, FONTS, LAYOUT, LOGO_FILES } from "./design.js";
import { fitBox, type Figure, type FigureFit, type FigureRenderer } from "./figure.js";
import { PAGE_NUMBER_SHAPE_NAME } from "./ooxml.js";
import { parseSvg, svgToGeomPoints } from "./svg-path.js";
import type { AssetResolver } from "./assets.js";
import { ImageResolver, type SlideImages } from "./images.js";
import type { ImageFit, ImageSource } from "./types.js";

type Slide = {
  addText: (text: string | TextRun[], options?: Record<string, unknown>) => unknown;
  addImage: (options: Record<string, unknown>) => unknown;
  addShape: (shapeName: string, options?: Record<string, unknown>) => unknown;
};

type ShapeType = {
  rect: string;
  line: string;
  roundRect: string;
  // pptxgenjs exposes the whole OOXML preset shape set — ellipse, triangle,
  // chevron and so on. The three above are the ones the helpers themselves
  // use; the index signature keeps the rest reachable from a custom slide
  // that legitimately needs one, instead of failing to typecheck over a shape
  // the renderer supports perfectly well.
  [name: string]: string;
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
  assets: AssetResolver;
  shapeType: ShapeType;
  /** Shared across the build. Without it, `addFigure` draws a placeholder. */
  figures?: FigureRenderer;
  /**
   * The build's picture resolver, bound to this slide. Without it, `addImage`
   * still places files, but records and warnings go nowhere.
   */
  images?: SlideImages;
}) {
  const { assets, shapeType, figures } = options;
  const images = options.images ?? new ImageResolver({ projectDir: assets.projectDir, warnings: [] }).forSlide(0);

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

    // The number is never passed in: the build rewrites this text box as a live
    // PowerPoint slide-number field, so it stays right when slides are reordered
    // here or dragged around in PowerPoint later. The "1" below is only the
    // fallback the field carries until the build fills in the real position.
    addFooter(slide: Slide, opts: { light?: boolean } = {}) {
      const light = opts.light ?? true;
      slide.addText([{ text: "-", options: { breakLine: true } }, { text: "1" }], {
        objectName: PAGE_NUMBER_SHAPE_NAME,
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
      const logoPath = assets.resolveLogo(opts.light ? LOGO_FILES.markDark : LOGO_FILES.markLight);
      if (!logoPath) return;
      slide.addImage({ path: logoPath, x: 9.3, y: 5.0, w: 0.33, h: 0.26 });
    },

    // Places the full wordmark, if the configured PNG exists in `assets/`.
    addWordmark(slide: Slide, opts: { light?: boolean; x?: number; y?: number; w?: number; h?: number } = {}) {
      const logoPath = assets.resolveLogo(opts.light ? LOGO_FILES.wordmarkDark : LOGO_FILES.wordmarkLight);
      if (!logoPath) return;
      slide.addImage({
        path: logoPath,
        x: opts.x ?? LAYOUT.LM,
        y: opts.y ?? 0.62,
        w: opts.w ?? 1.8,
        h: opts.h ?? 0.31
      });
    },

    addTextBlock(slide: Slide, runs: TextRun[], box: Box, style: TextOptions = {}) {
      slide.addText(runs, {
        ...box,
        fontSize: (style.fontSize as number | undefined) ?? 10,
        fontFace: (style.fontFace as string | undefined) ?? FONTS.sans,
        color: (style.color as string | undefined) ?? C.ink,
        lineSpacingMultiple: (style.lineSpacingMultiple as number | undefined) ?? LAYOUT.LS,
        valign: (style.valign as string | undefined) ?? "top",
        margin: (style.margin as number | undefined) ?? 0,
        ...style
      });
    },

    // `fill` and `color` move together. A card on a dark slide needs a dark
    // fill (C.inkSoft) AND light text (C.white); setting only the fill leaves
    // the heading and body in ink, which is unreadable. The colour is never
    // inferred from the fill: a caller who wants muted body text on a pale card
    // should be able to have it.
    addCard(
      slide: Slide,
      opts: Box & {
        heading: string;
        body?: string;
        accent?: string;
        fill?: string;
        /** Heading and body colour. Defaults to `C.ink`, which suits a light fill. */
        color?: string;
      }
    ) {
      const accent = stripHash(opts.accent ?? C.accent);
      const fill = stripHash(opts.fill ?? C.white);
      const color = stripHash(opts.color ?? C.ink);
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
        color,
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
          color,
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

    // Places a picture in `box`: an image declared in the brief's `## Images`
    // block (`{ image: "hero-city" }`, generated by `pptx-gen images`), or a
    // file in the deck folder (`{ path: "inputs/photo.jpg" }`). It fills the
    // box by default, cropping rather than distorting; `contain` shrinks the
    // box to the picture's shape instead. Returns the box actually used.
    //
    // A declared image with no file yet draws the captioned placeholder, so the
    // deck still builds and the slide still says what belongs there. Like a
    // figure, the result is a picture rather than editable parts: use it for
    // photos and illustrations, never for diagrams or text.
    async addImage(slide: Slide, source: ImageSource, box: Box, opts: { fit?: ImageFit; altText?: string } = {}) {
      const picture = await images.resolve(source);
      if (picture.status === "missing") {
        this.addImagePlaceholder(slide, { ...box, caption: picture.caption });
        return box;
      }
      const fit = images.fitFor(picture, opts.fit ?? "cover");
      const altText = opts.altText ?? picture.altText;
      const size = picture.size;
      if (fit === "cover" && size) {
        // pptxgenjs takes the picture's own proportions from w/h and the box
        // from `sizing`, and crops the difference. Passing the box as w/h
        // would tell it the shapes already match, and squash the picture.
        slide.addImage({
          path: picture.path,
          x: box.x,
          y: box.y,
          w: box.w,
          h: (box.w * size.pxHeight) / size.pxWidth,
          sizing: { type: "cover", w: box.w, h: box.h },
          altText
        });
        return box;
      }
      const placed = fit === "contain" && size ? fitBox(box, size.pxWidth, size.pxHeight, "contain") : box;
      slide.addImage({ path: picture.path, ...placed, altText });
      return placed;
    },

    // Renders an authored HTML figure and places it in `box`, preserving its
    // proportions. Reserve it for content that cannot be built from native
    // shapes — a UI mockup, a chart, a rendered document — because unlike every
    // other helper here the result is a picture, not editable parts.
    //
    // If the figure cannot be rendered (no browser installed, for instance) the
    // captioned placeholder is drawn instead, so the deck still builds and the
    // slide still says what belongs there. Returns the box actually used.
    async addFigure(slide: Slide, figure: Figure, box: Box, opts: { fit?: FigureFit } = {}): Promise<Box> {
      const result = await figures?.render(figure, { box: { w: box.w, h: box.h } });
      if (!result || result.status === "failed") {
        this.addImagePlaceholder(slide, { ...box, caption: figure.caption });
        return box;
      }
      const placed = fitBox(box, result.pxWidth, result.pxHeight, opts.fit ?? "contain");
      slide.addImage({ path: result.pngPath, ...placed, altText: figure.caption });
      return placed;
    },

    addArrow(
      slide: Slide,
      opts: {
        from: Point;
        to: Point;
        color?: string;
        width?: number;
        dashed?: boolean;
        beginArrowType?: ArrowType;
        endArrowType?: ArrowType;
      }
    ) {
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

    addConnector(
      slide: Slide,
      opts: {
        points: Point[];
        color?: string;
        width?: number;
        dashed?: boolean;
        endArrowType?: ArrowType;
      }
    ) {
      for (let index = 1; index < opts.points.length; index += 1) {
        this.addArrow(slide, {
          from: opts.points[index - 1],
          to: opts.points[index],
          color: opts.color,
          width: opts.width,
          dashed: opts.dashed,
          endArrowType: index === opts.points.length - 1 ? (opts.endArrowType ?? "triangle") : "none"
        });
      }
    },

    async addIcon(slide: Slide, icon: string, box: Box, opts: { color?: string; width?: number } = {}) {
      const iconPath = assets.resolveIcon(icon);
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

    addCodePanel(
      slide: Slide,
      opts: {
        code: string;
        x: number;
        y: number;
        w: number;
        maxH: number;
        fontFace?: string;
      }
    ) {
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
      slide.addText(
        lines.map((line, index) => {
          const trimmed = line.trimStart();
          const isComment = trimmed.startsWith("//") || trimmed.startsWith("#");
          return {
            text: line === "" ? " " : line,
            options: {
              color: isComment ? C.accent : C.white,
              breakLine: index < lines.length - 1
            }
          };
        }),
        {
          x: opts.x + 0.22,
          y: opts.y + 0.14,
          w: opts.w - 0.44,
          h: panelH - 0.28,
          fontSize: 10.5,
          fontFace: opts.fontFace ?? FONTS.mono,
          lineSpacingMultiple: 1.18,
          valign: "top",
          margin: 0
        }
      );
    }
  };
}

function stripHash(value: string): string {
  return value.replace(/^#/, "");
}
