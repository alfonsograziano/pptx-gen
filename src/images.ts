// Pictures at build time: turning `{ image: "hero" }` or `{ path: ... }` into
// bytes on a slide, or into a placeholder when the file is not there yet.
//
// This side never touches the network and never looks for an API key. Images
// are generated beforehand by `pptx-gen images` (image-gen.ts) and land in the
// deck's `inputs/` folder; the build only reads what is on disk, so the same
// inputs always produce the same deck (docs/decisions/0010).
import { existsSync, readFileSync } from "node:fs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { readPngSize } from "./html-shot.js";
import { parseImageRequests, type ImageRequest } from "./image-brief.js";
import { closestMatch } from "./workspace-config.js";
import type { BuildWarning, ImageFit, ImageRecord, ImageSource } from "./types.js";

export const BRIEF_FILE = "brief.md";
export const INPUTS_DIR = "inputs";

/**
 * The extensions a variant file may have, in lookup order. Generated files are
 * `.jpg` or `.png`; `.jpeg` covers a file someone dropped in by hand.
 */
export const VARIANT_EXTENSIONS = ["jpg", "jpeg", "png"] as const;

const CAPTION_LIMIT = 160;

export type ImageSize = { pxWidth: number; pxHeight: number };

export type ResolvedPicture =
  | {
      status: "found";
      path: string;
      bytes: Buffer;
      extension: string;
      /** Undefined for a format whose dimensions cannot be read. */
      size?: ImageSize;
      altText?: string;
    }
  | { status: "missing"; caption: string };

/** The resolver as one slide sees it, so every record carries its position. */
export type SlideImages = {
  resolve: (source: ImageSource) => Promise<ResolvedPicture>;
  /**
   * The fit to actually apply. `cover` and `contain` need the picture's
   * dimensions; without them this warns and falls back to `stretch`.
   */
  fitFor: (picture: Extract<ResolvedPicture, { status: "found" }>, fit: ImageFit) => ImageFit;
};

type BriefState = { requests: ImageRequest[] } | { error: string };

/** One per build: reads the brief once, and remembers every picture placed. */
export class ImageResolver {
  private readonly projectDir: string;
  private readonly warnings: BuildWarning[];
  private brief?: BriefState;
  private readonly placed: ImageRecord[] = [];
  private warnedInvalidBrief = false;
  private readonly warnedSize = new Set<string>();

  constructor(options: { projectDir: string; warnings: BuildWarning[] }) {
    this.projectDir = path.resolve(options.projectDir);
    this.warnings = options.warnings;
  }

  forSlide(slide: number): SlideImages {
    return {
      resolve: (source) => this.resolve(source, slide),
      fitFor: (picture, fit) => this.fitFor(picture, fit, slide)
    };
  }

  /** Every picture placed so far, in placement order. */
  records(): ImageRecord[] {
    return [...this.placed];
  }

  /** The brief's image requests; none when it is missing or malformed. */
  requests(): ImageRequest[] {
    const brief = this.loadBrief();
    return "requests" in brief ? brief.requests : [];
  }

  /** The `## Images` section of report.md, with links relative to `reportDir`. */
  formatReportSection(reportDir: string): string {
    const requests = this.requests();
    if (requests.length === 0 && this.placed.length === 0) return "- None";

    const link = (file: string) => path.relative(reportDir, file).split(path.sep).map(encodeURIComponent).join("/");
    const lines: string[] = [];
    for (const request of requests) {
      const uses = this.placed.filter((record) => record.id === request.id);
      const slides = [...new Set(uses.map((record) => record.slide))].join(", ");
      const intended = request.slide ? ` (meant for slide ${request.slide})` : "";
      let status: string;
      if (uses.length === 0) status = "declared in brief.md, but not placed in the deck";
      else if (uses.every((record) => record.status === "placed")) {
        status = `on slide ${slides}, variant ${request.pick} of ${request.variants}`;
      } else status = `placeholder on slide ${slides}: no file in ${INPUTS_DIR}/ yet`;
      lines.push(`- \`${request.id}\`${intended}: ${status}`);

      const files = Array.from({ length: request.variants }, (_unused, index) =>
        findVariantFile(this.projectDir, request.id, index + 1)
      );
      if (files.some(Boolean)) {
        const thumbnails = files.map((file, index) => {
          const label = `${request.id} ${index + 1}${index + 1 === request.pick ? " (picked)" : ""}`;
          return file ? `![${label}](${link(file)})` : `(variant ${index + 1} missing)`;
        });
        lines.push(`  ${thumbnails.join(" ")}`);
      }
    }
    for (const record of this.placed.filter((entry) => !entry.id)) {
      lines.push(`- ${record.path} (slide ${record.slide}): placed`);
    }
    return lines.join("\n");
  }

  private async resolve(source: ImageSource, slide: number): Promise<ResolvedPicture> {
    if (source.image === undefined) {
      // A path is the author saying "this file exists", so a miss stays the
      // hard error it has always been.
      const file = path.isAbsolute(source.path) ? source.path : path.resolve(this.projectDir, source.path);
      const bytes = await readFile(file);
      this.placed.push({ path: this.local(file), slide, status: "placed" });
      return { status: "found", path: file, bytes, extension: extensionOf(file), size: imageSize(bytes) };
    }

    const id = source.image;
    const brief = this.loadBrief();
    if ("error" in brief) {
      if (!this.warnedInvalidBrief) {
        this.warnedInvalidBrief = true;
        this.warnings.push({
          code: "image-brief-invalid",
          message: `The ## Images block in ${BRIEF_FILE} could not be read, so its images were drawn as placeholders. ${brief.error}`
        });
      }
      this.placed.push({ id, slide, status: "placeholder", reason: `${BRIEF_FILE} ## Images block is invalid` });
      return { status: "missing", caption: `Image needed: ${id}` };
    }

    const request = brief.requests.find((candidate) => candidate.id === id);
    if (!request) {
      const suggestion = closestMatch(
        id,
        brief.requests.map((candidate) => candidate.id)
      );
      throw new Error(
        `Image '${id}' is not declared in the ## Images block of ${path.join(this.projectDir, BRIEF_FILE)}` +
          `${suggestion ? ` — did you mean '${suggestion}'?` : ""}. Declare it there before placing it.`
      );
    }

    const file = findVariantFile(this.projectDir, id, request.pick);
    if (!file) {
      // Fixed wording, whatever the environment: the build never checks for an
      // API key, so the same inputs always produce the same report.
      this.warnings.push({
        code: "image-missing",
        message: `Image '${id}' has no file in ${INPUTS_DIR}/ yet (looked for ${id}-${request.pick}.${VARIANT_EXTENSIONS.join("|")}), so slide ${slide} shows a placeholder. Run \`pptx-gen images --project ${path.basename(this.projectDir)}\` to generate it, or put the file there yourself.`,
        slide,
        target: id
      });
      this.placed.push({ id, slide, status: "placeholder", reason: `no file in ${INPUTS_DIR}/ yet` });
      return { status: "missing", caption: placeholderCaption(request.description) };
    }

    const bytes = await readFile(file);
    this.placed.push({ id, path: this.local(file), slide, status: "placed" });
    return {
      status: "found",
      path: file,
      bytes,
      extension: extensionOf(file),
      size: imageSize(bytes),
      altText: request.description
    };
  }

  private fitFor(picture: Extract<ResolvedPicture, { status: "found" }>, fit: ImageFit, slide: number): ImageFit {
    if (fit === "stretch" || picture.size) return fit;
    if (!this.warnedSize.has(picture.path)) {
      this.warnedSize.add(picture.path);
      this.warnings.push({
        code: "image-size-unknown",
        message: `Could not read the pixel size of ${this.local(picture.path)}, so it was stretched to its box instead of using '${fit}'. Use a PNG or JPEG.`,
        slide,
        target: this.local(picture.path)
      });
    }
    return "stretch";
  }

  private loadBrief(): BriefState {
    if (this.brief) return this.brief;
    const briefPath = path.join(this.projectDir, BRIEF_FILE);
    if (!existsSync(briefPath)) {
      this.brief = { requests: [] };
      return this.brief;
    }
    try {
      this.brief = { requests: parseImageRequests(readFileSync(briefPath, "utf8"), briefPath) };
    } catch (error) {
      this.brief = { error: (error as Error).message };
    }
    return this.brief;
  }

  private local(file: string): string {
    const relative = path.relative(this.projectDir, file);
    return relative && !relative.startsWith("..") && !path.isAbsolute(relative) ? relative : file;
  }
}

/** The file for one variant of an image, whichever extension it has. */
function findVariantFile(projectDir: string, id: string, variant: number): string | undefined {
  for (const extension of VARIANT_EXTENSIONS) {
    const candidate = path.join(projectDir, INPUTS_DIR, `${id}-${variant}.${extension}`);
    if (existsSync(candidate)) return candidate;
  }
  return undefined;
}

function placeholderCaption(description: string): string {
  const text = description.replace(/\s+/g, " ").trim();
  const clipped = text.length > CAPTION_LIMIT ? `${text.slice(0, CAPTION_LIMIT - 1).trimEnd()}…` : text;
  return `Image needed: ${clipped}`;
}

function extensionOf(file: string): string {
  return (path.extname(file).slice(1) || "png").toLowerCase();
}

/** Pixel dimensions of a PNG or JPEG; undefined for anything else. */
export function imageSize(bytes: Buffer): ImageSize | undefined {
  try {
    return readPngSize(bytes);
  } catch {
    return readJpegSize(bytes);
  }
}

// Walk the JPEG segments to the first start-of-frame marker, which carries the
// height and width. No decoding: only headers are read.
function readJpegSize(bytes: Buffer): ImageSize | undefined {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = bytes[offset + 1];
    if (marker === 0xff) {
      offset += 1;
      continue;
    }
    // Markers with no length field.
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) {
      offset += 2;
      continue;
    }
    const isStartOfFrame = marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isStartOfFrame) {
      return { pxHeight: bytes.readUInt16BE(offset + 5), pxWidth: bytes.readUInt16BE(offset + 7) };
    }
    offset += 2 + bytes.readUInt16BE(offset + 2);
  }
  return undefined;
}

/**
 * The crop that makes a picture fill a box without distortion, as OOXML
 * `<a:srcRect>` insets in thousandths of a percent (100000 = the whole side).
 * Undefined when the shapes already match.
 */
export function coverCrop(
  box: { w: number; h: number },
  size: ImageSize
): { l: number; t: number; r: number; b: number } | undefined {
  if (box.w <= 0 || box.h <= 0 || size.pxWidth <= 0 || size.pxHeight <= 0) return undefined;
  const imageAspect = size.pxWidth / size.pxHeight;
  const boxAspect = box.w / box.h;
  if (Math.abs(imageAspect - boxAspect) / boxAspect < 0.001) return undefined;
  if (imageAspect > boxAspect) {
    const side = Math.round(((1 - boxAspect / imageAspect) / 2) * 100000);
    return { l: side, t: 0, r: side, b: 0 };
  }
  const side = Math.round(((1 - imageAspect / boxAspect) / 2) * 100000);
  return { l: 0, t: side, r: 0, b: side };
}
