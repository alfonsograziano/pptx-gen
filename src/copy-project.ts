#!/usr/bin/env node
/**
 * copy-project — pull all the text out of a project's built deck and put it on
 * the clipboard, ready to paste straight into Google Docs.
 *
 * Usage:
 *   npm run copy-project -- <project>            # rich text (HTML) onto the clipboard
 *   npm run copy-project -- <project> --text     # plain text instead
 *   npm run copy-project -- <project> --print    # also print to stdout, don't touch clipboard
 *
 * <project> is a folder name under projects/ (e.g. "custom-harness") or a path
 * to a project folder. The deck is read from <project>/output/*.pptx, preferring
 * deck.pptx when several exist.
 *
 * Only text is copied. Images, charts, and slide layout are not — paste those by
 * hand if you need them.
 */
import { spawn } from "node:child_process";
import { mkdtemp, readdir, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";
import { unescapeXml } from "./xml.js";
import { C } from "./design.js";

type Slide = { number: number; blocks: string[][] };

const HERE = path.dirname(fileURLToPath(import.meta.url));
const PROJECTS_ROOT = path.resolve(HERE, "..", "projects");

async function main(): Promise<void> {
  const { project, asText, printOnly } = parseArgs(process.argv.slice(2));
  const deckPath = await resolveDeckPath(project);

  const slides = await readSlides(deckPath);
  if (slides.length === 0) {
    throw new Error(`No slides with text found in ${deckPath}`);
  }

  const plain = toPlainText(slides);
  const html = toHtml(slides);
  const charCount = plain.length;

  if (printOnly) {
    process.stdout.write((asText ? plain : html) + "\n");
  } else if (asText) {
    await copyPlainText(plain);
  } else {
    await copyHtml(html);
  }

  const where = printOnly ? "printed" : "copied to clipboard";
  const kind = asText ? "plain text" : "rich text";
  console.log(
    `${kind} ${where}: ${slides.length} slide(s), ${charCount} characters from ${path.relative(process.cwd(), deckPath)}`
  );
  if (!printOnly) console.log("Paste it into Google Docs with Cmd+V.");
}

function parseArgs(argv: string[]): { project: string; asText: boolean; printOnly: boolean } {
  const positionals: string[] = [];
  let asText = false;
  let printOnly = false;
  for (const arg of argv) {
    if (arg === "--text" || arg === "-t") asText = true;
    else if (arg === "--print" || arg === "-p") printOnly = true;
    else if (arg.startsWith("-")) throw new Error(`Unknown option: ${arg}`);
    else positionals.push(arg);
  }
  const project = positionals[0];
  if (!project) {
    throw new Error("Usage: copy-project <project> [--text] [--print]");
  }
  return { project, asText, printOnly };
}

/** Find the deck PPTX for a project given a folder name or path. */
async function resolveDeckPath(project: string): Promise<string> {
  const candidates = project.includes(path.sep) || existsSync(project)
    ? [path.resolve(project)]
    : [path.join(PROJECTS_ROOT, project)];

  for (const projectDir of candidates) {
    const outputDir = path.join(projectDir, "output");
    if (!existsSync(outputDir)) continue;
    const pptxFiles = (await readdir(outputDir)).filter((file) => file.toLowerCase().endsWith(".pptx"));
    if (pptxFiles.length === 0) continue;
    // Prefer deck.pptx, otherwise take the first .pptx in the folder.
    const chosen = pptxFiles.find((file) => file === "deck.pptx") ?? pptxFiles.sort()[0];
    return path.join(outputDir, chosen);
  }

  throw new Error(
    `Could not find a built deck for '${project}'. Looked for a .pptx in ${candidates
      .map((dir) => path.relative(process.cwd(), path.join(dir, "output")))
      .join(", ")}. Build the deck first.`
  );
}

/** Read every slide's text, in presentation order, as ordered blocks of lines. */
async function readSlides(deckPath: string): Promise<Slide[]> {
  const pkg = await PptxPackage.load(deckPath);
  const entries = await getSlideEntries(pkg);
  const slides: Slide[] = [];

  for (const [index, entry] of entries.entries()) {
    const xml = await pkg.text(`ppt/slides/slide${entry.slideNumber}.xml`);
    const blocks = extractBlocks(xml);
    if (blocks.length > 0) slides.push({ number: index + 1, blocks });
  }

  return slides;
}

/**
 * Pull text out of a slide in reading order. Each text shape becomes a block,
 * each paragraph in that shape a line. `<a:br/>` becomes a line break inside the
 * paragraph. Empty shapes are dropped so the output stays clean.
 */
function extractBlocks(slideXml: string): string[][] {
  const shapes = slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
  const blocks: string[][] = [];

  for (const shape of shapes) {
    if (!shape.includes("<p:txBody>")) continue;
    const paragraphs = shape.match(/<a:p>[\s\S]*?<\/a:p>/g) ?? [];
    const lines = paragraphs
      .map(paragraphText)
      .filter((line) => line.trim().length > 0);
    if (lines.length > 0) blocks.push(lines);
  }

  return blocks;
}

/** Text of one paragraph, joining runs and turning <a:br/> into newlines. */
function paragraphText(paragraphXml: string): string {
  const tokens = paragraphXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>|<a:br\b[^>]*\/?>/g);
  let text = "";
  for (const token of tokens) {
    text += token[1] === undefined ? "\n" : unescapeXml(token[1]);
  }
  return text;
}

function toPlainText(slides: Slide[]): string {
  return slides
    .map((slide) => {
      const body = slide.blocks.map((lines) => lines.join("\n")).join("\n\n");
      return `--- Slide ${slide.number} ---\n${body}`;
    })
    .join("\n\n")
    .trim();
}

function toHtml(slides: Slide[]): string {
  const sections = slides
    .map((slide) => {
      const heading = `<h2 style="color:#${C.ink};">Slide ${slide.number}</h2>`;
      const paragraphs = slide.blocks
        .map((lines) => `<p>${lines.map(escapeHtml).join("<br>")}</p>`)
        .join("\n");
      return `${heading}\n${paragraphs}`;
    })
    .join("\n");
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>\n${sections}\n</body></html>`;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

async function copyPlainText(text: string): Promise<void> {
  await run("pbcopy", [], text);
}

/**
 * Put HTML on the macOS clipboard so Google Docs pastes it as formatted text.
 * pbcopy only handles plain text, so we write the HTML to a temp file and have
 * AppleScript read it onto the clipboard tagged as the HTML flavor.
 */
async function copyHtml(html: string): Promise<void> {
  if (process.platform !== "darwin") {
    throw new Error("Rich-text copy only works on macOS. Use --text for plain text on this platform.");
  }
  const dir = await mkdtemp(path.join(os.tmpdir(), "copy-project-"));
  const file = path.join(dir, "deck.html");
  await writeFile(file, html, "utf8");
  await run("osascript", [
    "-e",
    "on run argv",
    "-e",
    "set the clipboard to (read (POSIX file (item 1 of argv)) as «class HTML»)",
    "-e",
    "end run",
    file
  ]);
}

function run(command: string, args: string[], stdin?: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: [stdin === undefined ? "ignore" : "pipe", "ignore", "inherit"] });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with code ${code}`));
    });
    if (stdin !== undefined) {
      child.stdin!.end(stdin);
    }
  });
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
