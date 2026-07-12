import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const execFileAsync = promisify(execFile);

// Target long edge (px) for rasterized output, for crisp rendering when scaled up.
const TARGET_LONG_EDGE = 2000;

type Cache = string | undefined | null; // null = not looked up, undefined = looked up & missing
let rsvgCache: Cache = null;
let sofficeCache: Cache = null;

async function findExecutable(cache: Cache, candidates: string[]): Promise<[Cache, string | undefined]> {
  if (cache !== null) return [cache, cache ?? undefined];
  for (const candidate of candidates) {
    if (candidate.includes("/")) {
      try {
        await access(candidate);
        return [candidate, candidate];
      } catch {
        continue;
      }
    }
    try {
      const result = await execFileAsync("zsh", ["-lc", `command -v ${candidate}`]);
      const commandPath = result.stdout.trim();
      if (commandPath) return [commandPath, commandPath];
    } catch {
      continue;
    }
  }
  return [undefined, undefined];
}

async function findRsvg(): Promise<string | undefined> {
  const [cache, value] = await findExecutable(rsvgCache, [
    "rsvg-convert",
    "/opt/homebrew/bin/rsvg-convert",
    "/usr/local/bin/rsvg-convert"
  ]);
  rsvgCache = cache;
  return value;
}

async function findSoffice(): Promise<string | undefined> {
  const [cache, value] = await findExecutable(sofficeCache, [
    "soffice",
    "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice"
  ]);
  sofficeCache = cache;
  return value;
}

function intrinsicSize(svg: string): { w: number; h: number } | null {
  const open = svg.match(/<svg\b[^>]*>/);
  if (!open) return null;
  const tag = open[0];
  const wMatch = tag.match(/\bwidth="([\d.]+)"/);
  const hMatch = tag.match(/\bheight="([\d.]+)"/);
  const vbMatch = tag.match(/viewBox="([\d.\s-]+)"/);

  let w = wMatch ? Number.parseFloat(wMatch[1]) : undefined;
  let h = hMatch ? Number.parseFloat(hMatch[1]) : undefined;

  if ((!w || !h) && vbMatch) {
    const parts = vbMatch[1].trim().split(/\s+/).map(Number);
    if (parts.length === 4) {
      w = w ?? parts[2];
      h = h ?? parts[3];
    }
  }

  if (!w || !h) return null;
  return { w, h };
}

function scaledSize(svg: string): { w: number; h: number } {
  const size = intrinsicSize(svg);
  if (!size) return { w: TARGET_LONG_EDGE, h: TARGET_LONG_EDGE };
  const scale = Math.max(1, TARGET_LONG_EDGE / Math.max(size.w, size.h));
  return { w: Math.round(size.w * scale), h: Math.round(size.h * scale) };
}

// LibreOffice rasterizes at the root <svg> width/height, so scale those up.
// (Used only as a fallback; it does not preserve transparency.)
function withScaledDimensions(svg: string, w: number, h: number): string {
  const open = svg.match(/<svg\b[^>]*>/);
  if (!open) return svg;
  const tag = open[0];
  let newTag = tag.match(/\bwidth="[\d.]+"/)
    ? tag.replace(/\bwidth="[\d.]+"/, `width="${w}"`)
    : tag.replace(/<svg\b/, `<svg width="${w}"`);
  newTag = newTag.match(/\bheight="[\d.]+"/)
    ? newTag.replace(/\bheight="[\d.]+"/, `height="${h}"`)
    : newTag.replace(/<svg\b/, `<svg height="${h}"`);
  return svg.replace(tag, newTag);
}

async function rasterizeWithRsvg(rsvg: string, svg: string, w: number, h: number): Promise<Buffer | null> {
  const dir = await mkdtemp(path.join(os.tmpdir(), "nf-svg-"));
  const inPath = path.join(dir, "in.svg");
  const outPath = path.join(dir, "out.png");
  try {
    await writeFile(inPath, svg, "utf8");
    await execFileAsync(rsvg, ["-w", String(w), "-h", String(h), inPath, "-o", outPath]);
    return await readFile(outPath);
  } catch {
    return null;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

async function rasterizeWithSoffice(soffice: string, svg: string, w: number, h: number): Promise<Buffer | null> {
  const dir = await mkdtemp(path.join(os.tmpdir(), "nf-svg-"));
  const inPath = path.join(dir, "in.svg");
  const outPath = path.join(dir, "in.png");
  try {
    await writeFile(inPath, withScaledDimensions(svg, w, h), "utf8");
    await execFileAsync(soffice, ["--headless", "--convert-to", "png", "--outdir", dir, inPath]);
    return await readFile(outPath);
  } catch {
    return null;
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

/**
 * Rasterize an SVG string to a PNG buffer.
 *
 * PptxGenJS embeds SVG images with a "broken image" raster fallback, which is
 * what Google Slides (and older PowerPoint) show instead of the SVG. Shipping a
 * real PNG makes the diagram render everywhere.
 *
 * Prefers rsvg-convert because it preserves transparency (RGBA). Falls back to
 * LibreOffice (opaque white background) and finally to null, so callers can fall
 * back to embedding the raw SVG.
 */
export async function rasterizeSvgToPng(svg: string): Promise<Buffer | null> {
  const { w, h } = scaledSize(svg);

  const rsvg = await findRsvg();
  if (rsvg) {
    const png = await rasterizeWithRsvg(rsvg, svg, w, h);
    if (png) return png;
  }

  const soffice = await findSoffice();
  if (soffice) {
    const png = await rasterizeWithSoffice(soffice, svg, w, h);
    if (png) return png;
  }

  return null;
}
