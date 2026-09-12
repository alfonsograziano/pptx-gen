import { mkdtemp } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import pptxgenjs from "pptxgenjs";
import { C, FONTS, LAYOUT } from "./design.js";
import { createCustomSlideHelpers, type CustomSlideHelpers } from "./custom-slide-helpers.js";
import type { FigureRenderer } from "./figure.js";
import { ensureDir } from "./fs.js";
import type { AssetResolver } from "./assets.js";

type Pptx = {
  defineLayout: (layout: { name: string; width: number; height: number }) => void;
  layout: string;
  author: string;
  company: string;
  subject: string;
  title: string;
  theme: Record<string, unknown>;
  ShapeType: {
    rect: string;
    line: string;
    roundRect: string;
    // See the note on ShapeType in custom-slide-helpers.ts: pptxgenjs offers
    // the full preset shape set, so a custom slide can reach for any of them.
    [name: string]: string;
  };
  addSlide: () => Slide;
  writeFile: (options: { fileName: string }) => Promise<string>;
};

type Slide = {
  background: { color: string };
  addText: (text: string | Array<{ text: string; options?: Record<string, unknown> }>, options?: Record<string, unknown>) => unknown;
  addImage: (options: Record<string, unknown>) => unknown;
  addShape: (shapeName: string, options?: Record<string, unknown>) => unknown;
};

export type CustomSlideContext = {
  pptx: Pptx;
  slide: Slide;
  pageNum: number;
  projectDir: string;
  assetsDir: string;
  /** Resolves icon names, logo files, and project-relative paths. */
  assets: AssetResolver;
  design: {
    colors: typeof C;
    layout: typeof LAYOUT;
  };
  helpers: CustomSlideHelpers;
};

export type CustomSlideOptions = {
  name: string;
  background?: "light" | "dark" | { color: string };
  requiredFonts?: string[];
  /**
   * Optional variant-group id. Slides sharing a group are alternative
   * renderings of the same concept, listed together in the build report so the
   * reviewer can pick one. Variants of a group must be added consecutively.
   *
   * Deliberately not exposed on `CustomSlideContext`: a variant is a genuinely
   * different layout, not one layout branching on its own variant index.
   */
  group?: string;
  draw: (context: CustomSlideContext) => void | Promise<void>;
};

export class CustomSlide {
  readonly name: string;
  readonly background?: CustomSlideOptions["background"];
  readonly requiredFonts: string[];
  readonly group?: string;
  private readonly drawSlide: CustomSlideOptions["draw"];

  constructor(options: CustomSlideOptions) {
    this.name = options.name;
    this.background = options.background;
    this.group = options.group;
    this.requiredFonts = options.requiredFonts ?? [FONTS.sans];
    this.drawSlide = options.draw;
  }

  async draw(context: CustomSlideContext): Promise<void> {
    await this.drawSlide(context);
  }
}

export async function renderCustomSlideToPptx(options: {
  customSlide: CustomSlide;
  output: string;
  pageNum: number;
  assets: AssetResolver;
  title?: string;
  figures?: FigureRenderer;
}): Promise<void> {
  const pptx = createCustomPresentation(options.title);
  const slide = pptx.addSlide();
  applyBackground(slide, options.customSlide.background);
  const helpers = createCustomSlideHelpers({
    assets: options.assets,
    shapeType: pptx.ShapeType,
    figures: options.figures,
  });

  await options.customSlide.draw({
    pptx,
    slide,
    pageNum: options.pageNum,
    projectDir: options.assets.projectDir,
    assetsDir: options.assets.assetsDir,
    assets: options.assets,
    design: { colors: C, layout: LAYOUT },
    helpers
  });

  await ensureDir(path.dirname(options.output));
  await pptx.writeFile({ fileName: options.output });
}

export async function renderCustomSlidesToPptx(options: {
  customSlides: CustomSlide[];
  output: string;
  assets: AssetResolver;
  title?: string;
  figures?: FigureRenderer;
}): Promise<void> {
  const pptx = createCustomPresentation(options.title);
  const helpers = createCustomSlideHelpers({
    assets: options.assets,
    shapeType: pptx.ShapeType,
    figures: options.figures,
  });

  for (const [index, customSlide] of options.customSlides.entries()) {
    const slide = pptx.addSlide();
    applyBackground(slide, customSlide.background);
    await customSlide.draw({
      pptx,
      slide,
      pageNum: index + 1,
      projectDir: options.assets.projectDir,
      assetsDir: options.assets.assetsDir,
      assets: options.assets,
      design: { colors: C, layout: LAYOUT },
      helpers
    });
  }

  await ensureDir(path.dirname(options.output));
  await pptx.writeFile({ fileName: options.output });
}

export async function makeCustomSlideTempPath(): Promise<string> {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-gen-custom-slide-"));
  return path.join(dir, "slide.pptx");
}

function createCustomPresentation(title?: string): Pptx {
  const candidate = pptxgenjs as unknown as { default?: unknown };
  const PptxGenJS = (typeof pptxgenjs === "function" ? pptxgenjs : candidate.default) as new () => Pptx;
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: "WIDE_16_9", width: LAYOUT.width, height: LAYOUT.height });
  pptx.layout = "WIDE_16_9";
  pptx.author = "pptx-gen";
  pptx.company = "pptx-gen";
  pptx.subject = title ?? "Presentation";
  pptx.title = title ?? "Presentation";
  pptx.theme = {
    headFontFace: FONTS.sans,
    bodyFontFace: FONTS.sans,
    lang: "en-US"
  };
  return pptx;
}

function applyBackground(slide: Slide, background: CustomSlideOptions["background"]): void {
  if (!background || background === "light") {
    slide.background = { color: C.white };
  } else if (background === "dark") {
    slide.background = { color: C.ink };
  } else {
    slide.background = { color: background.color.replace(/^#/, "") };
  }
}
