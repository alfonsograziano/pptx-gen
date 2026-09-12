import { createHash } from "node:crypto";
import { readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";
import { C, FONTS } from "./design.js";
import { isFontInstalled } from "./fonts.js";
import { ensureDir } from "./fs.js";
import { chromeVersion, findChrome, screenshotHtml, type ShotFn } from "./html-shot.js";
import type { BuildWarning, FigureRecord } from "./types.js";

// A figure is HTML that the deck author writes, rendered by a headless browser
// at build time and placed on a slide as a picture.
//
// It exists for content that has no native representation at all: a product UI
// mockup, a chart, a rendered document. Everything that CAN be expressed as
// shapes, lines, text, and vector icons still should be, because those stay
// editable in PowerPoint and Google Slides and a figure does not.

/** Bump when the generated HTML or the browser flags change, to invalidate caches. */
const FIGURE_ENGINE_VERSION = "1";

const DEFAULT_VIEWPORT: FigureViewport = { width: 1280, height: 720 };
const DEFAULT_SCALE = 2;
/** Pixels per inch used by `viewport: "box"`; with scale 2 that is 256 ppi on the slide. */
const BOX_PX_PER_IN = 128;
const LARGE_PNG_BYTES = 2_000_000;

export type FigureViewport = { width: number; height: number };

/** The slide box a figure is headed for, in inches. */
export type FigureBox = { w: number; h: number };

export type FigureHint = { box?: FigureBox };

export type Figure = {
  /** Stable id; becomes the file name of the HTML and the PNG. */
  id: string;
  /**
   * What the figure shows, in one specific sentence. Shown verbatim in place of
   * the figure when it cannot be rendered, so write it for a reader.
   */
  caption: string;
  /** Inline HTML source. Use this or `htmlFile`. */
  html?: string;
  /** Path to an HTML file, resolved against the deck project directory. */
  htmlFile?: string;
  /** Layout canvas in CSS pixels. `"box"` derives it from the slide box. */
  viewport?: FigureViewport | "box";
  /** Device pixel ratio. Default 2, which is ~256 ppi on a full-width slide box. */
  scale?: number;
  /** Render on a transparent background instead of the page colour. */
  transparent?: boolean;
  /** Extra CSS appended after the injected brand block. */
  css?: string;
  /** Allow the page to reach the network. Off by default so builds stay reproducible. */
  allowNetwork?: boolean;
};

export type FigureFit = "contain" | "stretch";

export type Box = { x: number; y: number; w: number; h: number };

export type FigureResult =
  | { status: "rendered" | "cached"; id: string; htmlPath: string; pngPath: string; pxWidth: number; pxHeight: number }
  | { status: "failed"; id: string; htmlPath: string; reason: string };

export type FigureRendererOptions = {
  /** Directory the generated HTML and PNG are written to. */
  outputDir: string;
  /** Deck project directory, used to resolve `htmlFile`. */
  projectDir: string;
  warnings: BuildWarning[];
  /** Progress line for slow work; wired to the build's live output. */
  note?: (message: string) => void;
  /** Rasterizer override. Injected by tests so they need no browser. */
  shot?: ShotFn;
};

/**
 * Renders figures to PNGs, once each, and remembers what it produced.
 *
 * One instance is shared by a whole build so that a figure used on several
 * slides rasterizes once, and so that an absent browser is reported once rather
 * than once per figure.
 */
export class FigureRenderer {
  private readonly options: FigureRendererOptions;
  /** Rasterizations in flight or done, keyed by output path (i.e. by content). */
  private readonly inFlight = new Map<string, Promise<{ pxWidth: number; pxHeight: number }>>();
  private readonly records = new Map<string, FigureRecord>();
  private readonly usedFiles = new Set<string>();
  private readonly seenIds = new Map<string, string>();
  private warnedNoBrowser = false;
  private warnedFont = false;

  constructor(options: FigureRendererOptions) {
    this.options = options;
  }

  /**
   * Render `figure`, reusing the work if these exact pixels were already made.
   *
   * Environment problems (no browser, a browser that fails) come back as
   * `status: "failed"` with a warning, so the caller can fall back to a
   * placeholder and the deck still builds. Authoring mistakes — no source, a
   * missing file — throw, because only the author can fix those.
   */
  render(figure: Figure, hint: FigureHint = {}): Promise<FigureResult> {
    return this.renderOnce(figure, hint);
  }

  /** Every figure this build produced, for the build report. */
  results(): FigureRecord[] {
    return [...this.records.values()];
  }

  /**
   * Delete figure files this build did not produce.
   *
   * PNGs carry a content hash in their name, so editing a figure leaves the old
   * one behind; this is what stops the folder growing without bound. It must run
   * after every slide is drawn, because only then is the used-file set complete.
   */
  async prune(): Promise<void> {
    if (this.records.size === 0) return;
    const entries = await readdir(this.options.outputDir).catch(() => []);
    for (const entry of entries) {
      if (!entry.endsWith(".png") && !entry.endsWith(".html")) continue;
      const full = path.join(this.options.outputDir, entry);
      if (this.usedFiles.has(full)) continue;
      await rm(full, { force: true });
    }
  }

  private async renderOnce(figure: Figure, hint: FigureHint): Promise<FigureResult> {
    const source = await this.readSource(figure);
    const viewport = resolveViewport(figure, hint.box);
    const scale = figure.scale ?? DEFAULT_SCALE;
    const html = buildFigureHtml(figure, viewport, source);

    const htmlPath = path.join(this.options.outputDir, `${figure.id}.html`);
    await ensureDir(this.options.outputDir);
    await writeFile(htmlPath, html, "utf8");
    this.usedFiles.add(htmlPath);

    this.checkAspect(figure, viewport, hint.box);
    await this.checkFonts(figure);

    const shot = this.options.shot;
    const browser = shot ? undefined : await findChrome();
    if (!shot && !browser) {
      if (!this.warnedNoBrowser) {
        this.warnedNoBrowser = true;
        const override = process.env.CHROME_PATH?.trim();
        this.options.warnings.push({
          code: "figure-chrome-missing",
          message: override
            ? `CHROME_PATH points at '${override}', which is not an executable, so figures were replaced by captioned placeholders. The deck was still built.`
            : "No Chrome-family browser was found, so figures were replaced by captioned placeholders. The deck was still built. Install Google Chrome, or set CHROME_PATH, to render figures."
        });
      }
      return this.fail(figure, htmlPath, "no browser available");
    }

    const version = browser ? await chromeVersion(browser) : "injected";
    const hash = figureHash(html, viewport, scale, figure.transparent === true, version);
    const pngPath = path.join(this.options.outputDir, `${figure.id}.${hash}.png`);
    this.usedFiles.add(pngPath);
    this.checkIdReuse(figure, hash);

    // Keyed by output path, so the same figure drawn on several slides is
    // rasterized once while two different figures never share a result.
    const started = this.inFlight.get(pngPath);
    if (started) {
      const { pxWidth, pxHeight } = await started;
      return this.succeed("cached", figure, htmlPath, pngPath, pxWidth, pxHeight, 0);
    }

    const cached = await stat(pngPath).catch(() => undefined);
    if (cached?.isFile()) {
      const bytes = await readFile(pngPath);
      const dimensions = readPngDimensions(bytes);
      this.inFlight.set(pngPath, Promise.resolve(dimensions));
      this.options.note?.(`figure '${figure.id}': cached`);
      return this.succeed("cached", figure, htmlPath, pngPath, dimensions.pxWidth, dimensions.pxHeight, cached.size);
    }

    this.options.note?.(`figure '${figure.id}': rendering (${viewport.width}×${viewport.height} @${scale})`);
    const pending = (shot ?? screenshotHtml)({
      htmlPath,
      pngPath,
      width: viewport.width,
      height: viewport.height,
      scale,
      transparent: figure.transparent,
      allowNetwork: figure.allowNetwork
    });
    this.inFlight.set(pngPath, pending);
    try {
      const { pxWidth, pxHeight } = await pending;
      const size = (await stat(pngPath).catch(() => undefined))?.size ?? 0;
      return this.succeed("rendered", figure, htmlPath, pngPath, pxWidth, pxHeight, size);
    } catch (error) {
      this.inFlight.delete(pngPath);
      this.options.warnings.push({
        code: "figure-render-failed",
        message: `Figure '${figure.id}' could not be rendered: ${error instanceof Error ? error.message : String(error)}. A captioned placeholder was used instead; the generated HTML is at ${htmlPath}.`,
        target: figure.id
      });
      return this.fail(figure, htmlPath, "the browser failed to render it");
    }
  }

  private async readSource(figure: Figure): Promise<string> {
    if (figure.html !== undefined && figure.htmlFile !== undefined) {
      throw new Error(`Figure '${figure.id}' sets both 'html' and 'htmlFile'; use one.`);
    }
    if (figure.html !== undefined) return figure.html;
    if (figure.htmlFile === undefined) {
      throw new Error(`Figure '${figure.id}' has no source; set either 'html' or 'htmlFile'.`);
    }
    const sourcePath = path.resolve(this.options.projectDir, figure.htmlFile);
    try {
      return await readFile(sourcePath, "utf8");
    } catch {
      throw new Error(
        `Figure '${figure.id}' points at '${figure.htmlFile}', which was not found (looked in ${sourcePath}).`
      );
    }
  }

  private checkAspect(figure: Figure, viewport: FigureViewport, box?: FigureBox): void {
    if (box === undefined || figure.viewport === "box" || figure.viewport === undefined) return;
    const boxAspect = box.w / box.h;
    if (!Number.isFinite(boxAspect) || boxAspect <= 0) return;
    const aspect = viewport.width / viewport.height;
    if (Math.abs(aspect - boxAspect) / boxAspect <= 0.02) return;
    this.options.warnings.push({
      code: "figure-aspect-mismatch",
      message: `Figure '${figure.id}' is ${aspect.toFixed(2)}:1 but its slide box is ${boxAspect.toFixed(2)}:1, so it was shrunk to fit and leaves a gap. Give the viewport the box's ratio, or use viewport: "box".`,
      target: figure.id
    });
  }

  /**
   * Ids name files, so two different figures sharing one would overwrite each
   * other's HTML and leave the deck showing whichever ran last.
   */
  private checkIdReuse(figure: Figure, hash: string): void {
    const seen = this.seenIds.get(figure.id);
    if (seen === undefined) {
      this.seenIds.set(figure.id, hash);
      return;
    }
    if (seen === hash) return;
    this.options.warnings.push({
      code: "figure-id-reused",
      message: `Two different figures both use the id '${figure.id}', so they overwrite each other's files. Give each figure its own id.`,
      target: figure.id
    });
  }

  private async checkFonts(figure: Figure): Promise<void> {
    if (this.warnedFont) return;
    this.warnedFont = true;
    if (await isFontInstalled(FONTS.sans)) return;
    this.options.warnings.push({
      code: "figure-font-missing",
      message: `'${FONTS.sans}' is not installed on this machine, so figures render in a substitute font and will not match the deck. Run \`npm run install-fonts\`.`,
      target: figure.id
    });
  }

  private succeed(
    status: "rendered" | "cached",
    figure: Figure,
    htmlPath: string,
    pngPath: string,
    pxWidth: number,
    pxHeight: number,
    bytes: number
  ): FigureResult {
    if (bytes > LARGE_PNG_BYTES) {
      this.options.warnings.push({
        code: "figure-large-png",
        message: `Figure '${figure.id}' is ${(bytes / 1_000_000).toFixed(1)} MB, which bloats the deck. Lower its 'scale', or drop gradients and photographic detail.`,
        target: figure.id
      });
    }
    const existing = this.records.get(pngPath);
    this.records.set(pngPath, existing ?? { id: figure.id, status, htmlPath, pngPath, pxWidth, pxHeight });
    return { status, id: figure.id, htmlPath, pngPath, pxWidth, pxHeight };
  }

  private fail(figure: Figure, htmlPath: string, reason: string): FigureResult {
    this.records.set(htmlPath, { id: figure.id, status: "placeholder", htmlPath, reason });
    return { status: "failed", id: figure.id, htmlPath, reason };
  }
}

/** The layout canvas for a figure, in CSS pixels. */
export function resolveViewport(figure: Figure, box?: FigureBox): FigureViewport {
  const viewport = figure.viewport;
  if (viewport === undefined) return DEFAULT_VIEWPORT;
  if (viewport !== "box") return viewport;
  // "box" means "exactly the shape and size of the slide box", so the figure
  // fills it with no letterboxing and at a known, uniform pixel density.
  if (!box || !(box.w > 0) || !(box.h > 0)) return DEFAULT_VIEWPORT;
  return { width: Math.round(box.w * BOX_PX_PER_IN), height: Math.round(box.h * BOX_PX_PER_IN) };
}

/**
 * Fit a rendered figure inside a slide box.
 *
 * `contain` preserves the figure's proportions and centres it, because a
 * stretched UI mockup with squashed type is exactly the failure this feature
 * exists to avoid. `stretch` fills the box regardless.
 */
export function fitBox(box: Box, pxWidth: number, pxHeight: number, fit: FigureFit = "contain"): Box {
  if (fit === "stretch" || pxWidth <= 0 || pxHeight <= 0) return box;
  const aspect = pxWidth / pxHeight;
  const boxAspect = box.w / box.h;
  if (Math.abs(aspect - boxAspect) / boxAspect < 0.001) return box;
  if (aspect > boxAspect) {
    const h = box.w / aspect;
    return round({ x: box.x, y: box.y + (box.h - h) / 2, w: box.w, h });
  }
  const w = box.h * aspect;
  return round({ x: box.x + (box.w - w) / 2, y: box.y, w, h: box.h });
}

/**
 * Cache key for a rendered figure.
 *
 * The hash covers the finished HTML — so editing the design tokens invalidates
 * every figure automatically — plus everything else that changes the pixels.
 */
export function figureHash(
  html: string,
  viewport: FigureViewport,
  scale: number,
  transparent: boolean,
  browserVersion: string
): string {
  return createHash("sha1")
    .update([FIGURE_ENGINE_VERSION, html, viewport.width, viewport.height, scale, transparent, browserVersion].join(" "))
    .digest("hex")
    .slice(0, 12);
}

/**
 * Wrap authored markup in a page that is on-brand and exactly viewport-sized.
 *
 * The design tokens arrive as CSS custom properties, so figure markup names
 * colours the same way the rest of the deck does and follows a rebrand for free.
 */
export function buildFigureHtml(figure: Figure, viewport: FigureViewport, source: string): string {
  const style = `<style>\n${brandCss(viewport, figure.transparent === true)}${figure.css ? `\n${figure.css}\n` : ""}</style>`;

  if (/<html[\s>]/i.test(source)) {
    // A complete document: leave the author's structure alone and add the brand
    // block last so it still wins on specificity ties.
    if (/<\/head\s*>/i.test(source)) return source.replace(/<\/head\s*>/i, `${style}\n</head>`);
    return source.replace(/<html[^>]*>/i, (match) => `${match}\n<head>${style}</head>`);
  }

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>${escapeHtml(figure.caption)}</title>
${style}
</head>
<body>
${source}
</body>
</html>
`;
}

function brandCss(viewport: FigureViewport, transparent: boolean): string {
  const tokens = Object.entries(C).map(([name, value]) => `  --${cssVarName(name)}: #${value};`);
  return `:root {
${tokens.join("\n")}
  --font-sans: ${JSON.stringify(FONTS.sans)}, system-ui, -apple-system, "Segoe UI", "Helvetica Neue", Arial, sans-serif;
  --font-serif: ${JSON.stringify(FONTS.serif)}, Georgia, "Times New Roman", serif;
  --font-mono: ${JSON.stringify(FONTS.mono)}, ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
}
*, *::before, *::after { box-sizing: border-box; }
/* Pinned so that rem units mean the same thing regardless of browser defaults. */
html { font-size: 16px; }
html, body {
  margin: 0;
  padding: 0;
  width: ${viewport.width}px;
  height: ${viewport.height}px;
  overflow: hidden;
}
body {
  background: ${transparent ? "transparent" : "var(--white)"};
  color: var(--ink);
  font-family: var(--font-sans);
  -webkit-font-smoothing: antialiased;
}
img, svg { max-width: 100%; }
`;
}

// ink -> --ink, accent2 -> --accent-2, grey10 -> --grey-10, accentSoft -> --accent-soft
function cssVarName(name: string): string {
  return name
    .replace(/([a-z])([A-Z])/g, "$1-$2")
    .replace(/([a-zA-Z])(\d)/g, "$1-$2")
    .toLowerCase();
}

function readPngDimensions(bytes: Buffer): { pxWidth: number; pxHeight: number } {
  return { pxWidth: bytes.readUInt32BE(16), pxHeight: bytes.readUInt32BE(20) };
}

function round(box: Box): Box {
  const to3 = (value: number): number => Math.round(value * 1000) / 1000;
  return { x: to3(box.x), y: to3(box.y), w: to3(box.w), h: to3(box.h) };
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
