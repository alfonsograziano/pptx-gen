import { mkdtemp } from "node:fs/promises";
import { createRequire } from "node:module";
import os from "node:os";
import path from "node:path";
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
  addText: (
    text: string | Array<{ text: string; options?: Record<string, unknown> }>,
    options?: Record<string, unknown>
  ) => unknown;
  addImage: (options: Record<string, unknown>) => unknown;
  addShape: (shapeName: string, options?: Record<string, unknown>) => unknown;
  addNotes: (notes: string) => unknown;
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
  /**
   * Speaker notes, one paragraph per line. Replaces anything `draw` passed to
   * pptxgenjs's own `slide.addNotes`, which also survives into the deck.
   */
  notes?: string;
  draw: (context: CustomSlideContext) => void | Promise<void>;
};

export class CustomSlide {
  readonly name: string;
  readonly background?: CustomSlideOptions["background"];
  readonly requiredFonts: string[];
  readonly group?: string;
  readonly notes?: string;
  private readonly drawSlide: CustomSlideOptions["draw"];

  constructor(options: CustomSlideOptions) {
    this.name = options.name;
    this.background = options.background;
    this.group = options.group;
    this.notes = options.notes;
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
    figures: options.figures
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
    figures: options.figures
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
  const PptxGenJS = loadPptxGenJS();
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

type PptxCtor = new () => Pptx;

let pptxCtor: PptxCtor | undefined;

/**
 * Load pptxgenjs through `require`, on first use, and keep it that way.
 *
 * pptxgenjs publishes `dist/pptxgen.es.js` under its `import` condition: ESM
 * source in a `.js` file, inside a package with no `"type": "module"`. Node
 * therefore has to load that file as CommonJS and fall back to `require(esm)`,
 * and `require(esm)` is illegal inside an import cycle. `bin/pptx-gen.mjs`
 * registers the tsx ESM loader in-process and then imports the CLI, so the
 * whole engine is loaded from inside a live cycle — and on Node 23 and 24 the
 * CLI died before building anything: "Cannot require() ES Module
 * .../pptxgen.es.js in a cycle". Node 22 and below, and Node 25, tolerate it.
 *
 * Taking the `require` condition instead gets `dist/pptxgen.cjs.js`, a real
 * CommonJS file, so no `require(esm)` happens and there is no cycle to be
 * inside of. Deferring it to first use also keeps the library out of the module
 * graph entirely for decks built only from templates.
 *
 * Do not "modernise" this back into a top-level or dynamic `import` of
 * pptxgenjs: that is exactly the shape that broke, and the package's `engines`
 * field promises Node >= 20.
 */
function loadPptxGenJS(): PptxCtor {
  if (pptxCtor) return pptxCtor;

  const require = createRequire(import.meta.url);
  const loaded = require("pptxgenjs") as unknown;
  // Depending on how the CJS build was bundled the constructor is either the
  // export itself or its `default`, so take whichever is callable.
  const candidate = typeof loaded === "function" ? loaded : (loaded as { default?: unknown })?.default;
  if (typeof candidate !== "function") {
    throw new Error("pptxgenjs did not export a constructor; the installed version may be incompatible.");
  }

  pptxCtor = candidate as PptxCtor;
  return pptxCtor;
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
