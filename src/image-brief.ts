// The words behind a generated image: what the deck's brief asks for, and how
// the workspace says images should look.
//
// Two markdown files feed image generation, and both belong to people:
//
//   brief.md      a `## Images` section holding one fenced yaml block, one
//                 entry per image the deck needs.
//   customize.md  an optional `## Image generation` section: style prose, a
//                 list of example images, and a palette switch.
//
// Everything here is a pure function of those strings, so the same brief and
// customizations always compose the same prompt. The network lives in
// image-gen.ts and the build-time lookup in images.ts.
import os from "node:os";
import path from "node:path";
import { parse as parseYaml } from "yaml";
import { closestMatch } from "./workspace-config.js";

export type ImageAspect = "landscape" | "portrait" | "square" | "wide";

/** One image declared in a brief's `## Images` block. */
export type ImageRequest = {
  /** Kebab-case, unique within the brief. Names the files: `inputs/<id>-<n>.jpg`. */
  id: string;
  /** What the image shows. The heart of the prompt, and the placeholder's caption. */
  description: string;
  /** The slide it is meant for. Informational: it appears in the build report. */
  slide?: number;
  /** How many takes to generate, 1–4. */
  variants: number;
  /** Which take the deck uses, 1-based. */
  pick: number;
  aspect: ImageAspect;
  /** An explicit `WxH`, instead of an aspect preset. */
  size?: string;
  /** Extra guidance for this image only, on top of customize.md's. */
  style?: string;
  /** Project-relative example images for this image only. */
  references: string[];
  /** Overrides customize.md's palette switch for this image. */
  palette?: boolean;
};

/** customize.md's `## Image generation` section. */
export type ImageStyle = {
  guidance: string;
  /** Absolute paths of example images sent along as style references. */
  examples: string[];
  /** Whether prompts carry the brand-colour hint. */
  palette: boolean;
};

export const NO_IMAGE_STYLE: ImageStyle = { guidance: "", examples: [], palette: true };

const MAX_VARIANTS = 4;

const REQUEST_KEYS = [
  "id",
  "description",
  "slide",
  "variants",
  "pick",
  "aspect",
  "size",
  "style",
  "references",
  "palette"
] as const;

const ASPECTS: readonly ImageAspect[] = ["landscape", "portrait", "square", "wide"];

// Sizes every gpt-image model accepts, except `wide`, which needs a model with
// flexible sizes. Placement crops to the slide's box anyway, so the preset only
// has to be close.
const ASPECT_SIZES: Record<ImageAspect, string> = {
  landscape: "1536x1024",
  portrait: "1024x1536",
  square: "1024x1024",
  wide: "1536x864"
};

const ASPECT_WORDS: Record<ImageAspect, string> = {
  landscape: "landscape (3:2)",
  portrait: "portrait (2:3)",
  square: "square",
  wide: "wide 16:9"
};

const ID_PATTERN = /^[a-z0-9][a-z0-9-]*$/;

class ImageBriefError extends Error {
  constructor(source: string, message: string) {
    super(`${source}: ${message}`);
    this.name = "ImageBriefError";
  }
}

/**
 * Read the image requests out of a brief.
 *
 * No `## Images` section, or one with no yaml fence, is simply zero requests:
 * most decks have no generated imagery. Anything malformed inside the block is
 * an authoring error and throws, naming the entry and the key.
 */
export function parseImageRequests(markdown: string, source = "brief.md"): ImageRequest[] {
  const block = findImagesBlock(markdown);
  if (!block?.fence) return [];

  let parsed: unknown;
  try {
    parsed = parseYaml(markdown.slice(block.fence.contentStart, block.fence.contentEnd));
  } catch (error) {
    throw new ImageBriefError(source, `the ## Images yaml block does not parse: ${(error as Error).message}`);
  }
  if (parsed === null || parsed === undefined) return [];
  if (!Array.isArray(parsed)) {
    throw new ImageBriefError(source, "the ## Images yaml block must be a list, one `- id: ...` entry per image.");
  }

  const requests = parsed.map((entry, index) => validateRequest(entry, index, source));
  const seen = new Set<string>();
  for (const request of requests) {
    if (seen.has(request.id)) throw new ImageBriefError(source, `image id '${request.id}' is declared twice.`);
    seen.add(request.id);
  }
  return requests;
}

function validateRequest(entry: unknown, index: number, source: string): ImageRequest {
  const where = `image ${index + 1}`;
  if (typeof entry !== "object" || entry === null || Array.isArray(entry)) {
    throw new ImageBriefError(source, `${where} must be a mapping of keys, like \`- id: hero\`.`);
  }
  const record = entry as Record<string, unknown>;
  const label = typeof record.id === "string" ? `image '${record.id}'` : where;
  const fail = (message: string): never => {
    throw new ImageBriefError(source, `${label}: ${message}`);
  };

  for (const key of Object.keys(record)) {
    if ((REQUEST_KEYS as readonly string[]).includes(key)) continue;
    const suggestion = closestMatch(key, REQUEST_KEYS);
    fail(
      `unknown key "${key}"${suggestion ? ` — did you mean "${suggestion}"?` : ""} Valid keys: ${REQUEST_KEYS.join(", ")}.`
    );
  }

  const id = record.id;
  if (typeof id !== "string" || !ID_PATTERN.test(id)) {
    fail(`"id" must be kebab-case (lowercase letters, digits and dashes), got ${describe(id)}.`);
  }
  const description = record.description;
  if (typeof description !== "string" || description.trim() === "") {
    fail(`"description" is required: say what the image shows.`);
  }

  const slide = optionalInteger(record.slide, "slide", 1, Number.POSITIVE_INFINITY, fail);
  const variants = optionalInteger(record.variants, "variants", 1, MAX_VARIANTS, fail) ?? 1;
  const pick = optionalInteger(record.pick, "pick", 1, Number.POSITIVE_INFINITY, fail) ?? 1;
  if (pick > variants) fail(`"pick" is ${pick}, but only ${variants} variant(s) are generated.`);

  if (record.aspect !== undefined && record.size !== undefined) {
    fail(`give "aspect" or "size", not both.`);
  }
  const aspect = record.aspect ?? "landscape";
  if (typeof aspect !== "string" || !(ASPECTS as readonly string[]).includes(aspect)) {
    fail(`"aspect" must be one of ${ASPECTS.join(", ")}, got ${describe(aspect)}.`);
  }
  const size = record.size;
  if (size !== undefined && (typeof size !== "string" || !isValidSize(size))) {
    fail(`"size" must be WIDTHxHEIGHT with both sides multiples of 16, like 1920x1088; got ${describe(size)}.`);
  }

  const style = record.style;
  if (style !== undefined && typeof style !== "string") fail(`"style" must be text, got ${describe(style)}.`);

  const rawReferences = record.references ?? [];
  const references = typeof rawReferences === "string" ? [rawReferences] : rawReferences;
  if (!Array.isArray(references) || references.some((item) => typeof item !== "string" || item.trim() === "")) {
    fail(`"references" must be a path or a list of paths, relative to the deck folder.`);
  }

  const palette = record.palette;
  if (palette !== undefined && typeof palette !== "boolean") {
    fail(`"palette" must be true or false, got ${describe(palette)}.`);
  }

  return {
    id: id as string,
    description: (description as string).trim(),
    slide,
    variants,
    pick,
    aspect: aspect as ImageAspect,
    size: size as string | undefined,
    style: (style as string | undefined)?.trim() || undefined,
    references: (references as string[]).map((item) => item.trim()),
    palette: palette as boolean | undefined
  };
}

function optionalInteger(
  value: unknown,
  key: string,
  min: number,
  max: number,
  fail: (message: string) => never
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "number" || !Number.isInteger(value) || value < min || value > max) {
    const range = Number.isFinite(max) ? `between ${min} and ${max}` : `${min} or more`;
    fail(`"${key}" must be a whole number ${range}, got ${describe(value)}.`);
  }
  return value as number;
}

function isValidSize(size: string): boolean {
  const match = size.match(/^(\d+)x(\d+)$/);
  if (!match) return false;
  const [width, height] = [Number(match[1]), Number(match[2])];
  return width > 0 && height > 0 && width % 16 === 0 && height % 16 === 0;
}

function describe(value: unknown): string {
  if (value === undefined) return "nothing";
  if (value === null) return "null";
  if (Array.isArray(value)) return "a list";
  if (typeof value === "string") return `"${value}"`;
  return `${typeof value} ${String(value)}`;
}

/**
 * Read customize.md's `## Image generation` section.
 *
 * The section is prose for a person first: an `Examples:` label introduces a
 * bullet list of image paths (relative to the workspace root), and a
 * `Palette: off` line drops the brand-colour hint. Every other line is style
 * guidance, passed to the model as written. HTML comments are ignored, so the
 * section can carry its own instructions.
 */
export function parseImageStyle(markdown: string, workspaceRoot: string): ImageStyle {
  const text = stripComments(markdown);
  const section = sectionBody(text, /^##\s+image generation\s*$/im);
  if (section === undefined) return { ...NO_IMAGE_STYLE };

  const guidance: string[] = [];
  const examples: string[] = [];
  let palette = true;
  let inExamples = false;

  for (const line of section.split(/\r?\n/)) {
    const trimmed = line.trim();
    const paletteMatch = trimmed.match(/^palette\s*:\s*(\S+)/i);
    if (paletteMatch) {
      palette = !/^(off|no|false|none)$/i.test(paletteMatch[1]);
      inExamples = false;
      continue;
    }
    if (/^(examples?|references?)\s*:?\s*$/i.test(trimmed) || /^#{3,}\s+(examples?|references?)\s*$/i.test(trimmed)) {
      inExamples = true;
      continue;
    }
    const bullet = trimmed.match(/^[-*+]\s+(.+)$/);
    if (inExamples && bullet) {
      examples.push(resolveStylePath(bulletPath(bullet[1]), workspaceRoot));
      continue;
    }
    if (inExamples && trimmed === "") continue;
    inExamples = false;
    guidance.push(line);
  }

  return { guidance: unwrap(tidy(guidance.join("\n"))), examples, palette };
}

// Markdown wraps prose at whatever width the author's editor used; join those
// soft breaks back into paragraphs, but keep list items on their own lines.
function unwrap(text: string): string {
  return text.replace(/([^\n])\n(?![-*+]\s|\d+\.\s|\n)/g, "$1 ");
}

// "`assets/a.jpg`", "![alt](assets/a.jpg)" and "assets/a.jpg — the good one"
// all name the same file.
function bulletPath(value: string): string {
  const link = value.match(/\]\(([^)]+)\)/);
  if (link) return link[1].trim();
  const code = value.match(/`([^`]+)`/);
  if (code) return code[1].trim();
  return value.split(/\s+(?:#|—|--)\s+/)[0].trim();
}

function resolveStylePath(value: string, workspaceRoot: string): string {
  if (value === "~") return os.homedir();
  if (value.startsWith("~/")) return path.join(os.homedir(), value.slice(2));
  return path.resolve(workspaceRoot, value);
}

/** The brand colours worth hinting at: the dark tone and the accents, deduplicated. */
export function brandPalette(colors: Record<string, string>): string[] {
  const picked: string[] = [];
  for (const name of ["ink", "accent", "accent2", "accent3"]) {
    const hex = colors[name]?.replace(/^#/, "").toUpperCase();
    if (hex && !picked.some((entry) => entry.startsWith(`#${hex}`))) picked.push(`#${hex} (${name})`);
  }
  return picked;
}

/**
 * The full prompt for one image, in a fixed order.
 *
 * Deterministic on purpose: the prompt is hashed to decide whether an image
 * needs regenerating, so the same brief and customizations must always yield
 * the same text.
 */
export function composeImagePrompt(
  request: ImageRequest,
  style: ImageStyle,
  options: { palette: string[]; hasReferences: boolean }
): string {
  const parts = [request.description];
  if (request.style) parts.push(request.style);
  if (style.guidance) parts.push(`Style: ${style.guidance}`);
  if ((request.palette ?? style.palette) && options.palette.length > 0) {
    parts.push(`Brand colours: where it suits the scene, lean on ${options.palette.join(", ")}.`);
  }
  if (options.hasReferences) {
    parts.push(
      "The attached images are style references only: match their look, palette and mood, but do not copy their subjects or composition."
    );
  }
  const shape = request.size ?? ASPECT_WORDS[request.aspect];
  parts.push(`Format: a ${shape} image for a presentation slide. No text, lettering, captions, logos or watermarks.`);
  return parts.join("\n\n");
}

/** The `WxH` to request for an image. */
export function sizeFor(request: ImageRequest): string {
  return request.size ?? ASPECT_SIZES[request.aspect];
}

const NOTES_START = "<!-- pptx-gen:image-notes:start -->";
const NOTES_END = "<!-- pptx-gen:image-notes:end -->";
const GENERAL = "*";

/**
 * Update the notes `pptx-gen images` keeps in the brief, and return the new
 * markdown.
 *
 * The notes live in one managed block right after the `## Images` yaml fence,
 * so a reader sees what went wrong next to what was asked for. Only the ids in
 * `processed` are rewritten — a run limited with `--only` leaves the others'
 * notes alone — and the general note (no key, say) is replaced on every run.
 * The block disappears once it has nothing to say, and nothing outside it is
 * ever touched.
 */
export function updateImageNotes(
  markdown: string,
  update: { processed: string[]; notes: Map<string, string>; general?: string }
): string {
  const existing = findNotesBlock(markdown);
  const notes = existing ? parseNotes(markdown.slice(existing.start, existing.end)) : new Map<string, string>();

  notes.delete(GENERAL);
  for (const id of update.processed) notes.delete(id);
  for (const [id, note] of update.notes) notes.set(id, note);
  if (update.general) notes.set(GENERAL, update.general);

  const rendered = notes.size > 0 ? renderNotes(notes) : "";

  if (existing) {
    if (rendered) return markdown.slice(0, existing.start) + rendered + markdown.slice(existing.end);
    // Take the blank line the block was inserted with, so removing it restores
    // the brief exactly.
    const start = markdown.slice(0, existing.start).endsWith("\n\n") ? existing.start - 1 : existing.start;
    const end = markdown.slice(existing.end).startsWith("\n") ? existing.end + 1 : existing.end;
    return markdown.slice(0, start) + markdown.slice(end);
  }
  if (!rendered) return markdown;

  const block = findImagesBlock(markdown);
  if (!block?.fence) return markdown;
  const at = block.fence.end;
  const after = markdown.slice(at);
  const lineBreak = after.startsWith("\n") ? "" : "\n";
  return `${markdown.slice(0, at)}\n\n${rendered}${lineBreak}${after}`;
}

function renderNotes(notes: Map<string, string>): string {
  const lines = [
    NOTES_START,
    "> **Image generation notes** — written by `pptx-gen images`. Until each is resolved, the slide shows a grey placeholder.",
    ">"
  ];
  const general = notes.get(GENERAL);
  if (general) lines.push(`> - ${oneLine(general)}`);
  for (const [id, note] of [...notes].filter(([key]) => key !== GENERAL).sort(([a], [b]) => a.localeCompare(b))) {
    lines.push(`> - \`${id}\`: ${oneLine(note)}`);
  }
  lines.push(NOTES_END);
  return lines.join("\n");
}

function parseNotes(block: string): Map<string, string> {
  const notes = new Map<string, string>();
  for (const line of block.split(/\r?\n/)) {
    const withId = line.match(/^>\s*-\s*`([^`]+)`:\s*(.*)$/);
    if (withId) {
      notes.set(withId[1], withId[2]);
      continue;
    }
    const general = line.match(/^>\s*-\s*(.+)$/);
    if (general) notes.set(GENERAL, general[1]);
  }
  return notes;
}

function findNotesBlock(markdown: string): { start: number; end: number } | undefined {
  const start = markdown.indexOf(NOTES_START);
  if (start === -1) return undefined;
  const endMarker = markdown.indexOf(NOTES_END, start);
  if (endMarker === -1) return undefined;
  return { start, end: endMarker + NOTES_END.length };
}

type ImagesBlock = {
  fence?: { start: number; end: number; contentStart: number; contentEnd: number };
};

/**
 * Locate the `## Images` section and its first yaml fence in the ORIGINAL text.
 *
 * Offsets have to be into the brief as written, because the notes are spliced
 * back into it. So instead of stripping HTML comments first, matches that fall
 * inside one are skipped — which is also what lets a commented-out example
 * block sit in the section without being read.
 */
function findImagesBlock(markdown: string): ImagesBlock | undefined {
  const comments = [...markdown.matchAll(/<!--[\s\S]*?-->/g)].map((match) => [
    match.index,
    match.index + match[0].length
  ]);
  const inComment = (offset: number) => comments.some(([start, end]) => offset >= start && offset < end);
  const firstOutside = (pattern: RegExp, from: number) => {
    pattern.lastIndex = from;
    for (let match = pattern.exec(markdown); match; match = pattern.exec(markdown)) {
      if (!inComment(match.index)) return match;
    }
    return undefined;
  };

  const heading = firstOutside(/^##\s+images\s*$/gim, 0);
  if (!heading) return undefined;
  const bodyStart = heading.index + heading[0].length;
  const nextHeading = firstOutside(/^#{1,2}\s+\S/gm, bodyStart);
  const sectionEnd = nextHeading ? nextHeading.index : markdown.length;

  const open = firstOutside(/^(`{3,}|~{3,})[ \t]*(?:yaml|yml)[ \t]*\r?$/gim, bodyStart);
  if (!open || open.index >= sectionEnd) return {};
  const contentStart = open.index + open[0].length + 1;
  const close = new RegExp(`^${open[1][0] === "`" ? "`" : "~"}{${open[1].length},}[ \\t]*\\r?$`, "gm");
  close.lastIndex = contentStart;
  const closing = close.exec(markdown);
  if (!closing) return {};
  return {
    fence: {
      start: open.index,
      end: closing.index + closing[0].length,
      contentStart: Math.min(contentStart, closing.index),
      contentEnd: closing.index
    }
  };
}

function stripComments(markdown: string): string {
  return markdown.replace(/<!--[\s\S]*?-->/g, "");
}

function sectionBody(markdown: string, heading: RegExp): string | undefined {
  const match = markdown.match(heading);
  if (match?.index === undefined) return undefined;
  const rest = markdown.slice(match.index + match[0].length);
  const next = rest.search(/^#{1,2}\s+\S/m);
  return next === -1 ? rest : rest.slice(0, next);
}

function tidy(text: string): string {
  return text
    .split(/\r?\n/)
    .map((line) => line.trimEnd())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function oneLine(text: string): string {
  return text.replace(/\s*\r?\n\s*/g, " ").trim();
}
