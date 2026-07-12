/**
 * Ensure every font the template library needs is available.
 *
 * The script scans all templates, lists the typefaces they require, and reports
 * two kinds of coverage for each:
 *
 *  - "embedded": the font binary travels inside the template's PPTX, so any deck
 *    built from it is self-contained and renders correctly even on a machine
 *    that has never seen the font. This is what the generator relies on.
 *  - "system":  the font is installed on this machine, which matters for editing
 *    the PPTX natively in PowerPoint / Keynote and for local LibreOffice renders.
 *
 * The Nearform brand families (Inter, Bitter, Lexend) are installed from Google
 * Fonts. Families that are not in the catalogue and not embedded anywhere are
 * reported as a hard gap so you can decide on a fallback.
 *
 * IMPORTANT — why we install STATIC instances, not variable fonts:
 * Google ships these families as variable fonts (one `Family[wght].ttf` file).
 * LibreOffice on macOS cannot resolve those variable files: it fails to match
 * the requested family and silently substitutes a default sans-serif. That is
 * why the PDF / screenshot renders showed the wrong font (e.g. Bitter, a serif,
 * came out sans-serif) even though Google Slides — which reads the font embedded
 * in the PPTX — looked correct. The variable Bitter file also defaults to its
 * "Thin" instance, which compounds the mismatch.
 *
 * The fix: download per-weight STATIC instances via the Google Fonts CSS2 API
 * (the same files Google serves to browsers). Each static file registers with a
 * clean family name ("Bitter", "Bitter Medium", "Inter SemiBold", ...) that
 * LibreOffice resolves reliably, so local renders match what clients see. We
 * also delete any previously installed variable brand files so the broken ones
 * can no longer shadow the good static ones.
 */
import { access, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";
import { extractFonts } from "./ooxml.js";

const execFileAsync = promisify(execFile);

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE_ROOT = path.resolve(HERE, "..", "templates");
const USER_FONT_DIR = path.join(os.homedir(), "Library", "Fonts");

// Fonts that never need installing: they ship with every OS or are handled by
// the renderer's own substitution.
const GENERIC_FONTS = new Set(["Arial", "Calibri", "Aptos", "Helvetica", "Times New Roman"]);

// Brand font families and the static weights to install for each. We install a
// generous set so any deck's chosen weights resolve, including the bold variant
// used when a run sets b="1". Each weight becomes its own static .ttf file.
//
// We deliberately do NOT use the variable `Family[wght].ttf` files here: see the
// file header for why LibreOffice cannot render them.
const BRAND_FONTS: Record<string, { weights: number[]; italicWeights: number[] }> = {
  Inter: { weights: [300, 400, 500, 600, 700, 900], italicWeights: [400, 700] },
  Bitter: { weights: [400, 500, 700], italicWeights: [400, 500, 700] },
  Lexend: { weights: [400, 500, 700], italicWeights: [] }
};

// Weight number -> the word Google Fonts uses, for human-readable file names.
const WEIGHT_NAMES: Record<number, string> = {
  100: "Thin", 200: "ExtraLight", 300: "Light", 400: "Regular",
  500: "Medium", 600: "SemiBold", 700: "Bold", 800: "ExtraBold", 900: "Black"
};

// The variable brand files the old version of this script installed. They make
// LibreOffice substitute a wrong (sans-serif) font, so we remove them before
// installing the static replacements.
const STALE_VARIABLE_FILES = [
  "Inter.ttf", "Inter-Italic.ttf",
  "Bitter.ttf", "Bitter-Italic.ttf",
  "Lexend.ttf", "Lexend-Italic.ttf"
];

// A deliberately ancient User-Agent so the Google Fonts CSS2 API serves plain
// TrueType (.ttf) files. A modern UA gets woff2 (macOS will not install it); a
// too-specific old UA (e.g. an IE6 string) gets a degraded single-face EOT-style
// response. The bare "Mozilla/4.0" reliably returns one .ttf per weight.
const LEGACY_USER_AGENT = "Mozilla/4.0";

type StaticFace = { family: string; weight: number; italic: boolean; url: string };

const WEIGHT_WORDS = new Set([
  "thin", "extralight", "ultralight", "light", "regular", "medium",
  "semibold", "demibold", "bold", "extrabold", "ultrabold", "black", "heavy", "italic"
]);

function baseFamily(font: string): string {
  const words = font.trim().split(/\s+/);
  while (words.length > 1 && WEIGHT_WORDS.has(words[words.length - 1].toLowerCase())) words.pop();
  return words.join(" ");
}

type FontStatus = {
  typeface: string;
  family: string;
  embedded: boolean;
  system: boolean;
};

async function main(): Promise<void> {
  const templates = (await readdir(TEMPLATE_ROOT, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();

  // 1. Gather required typefaces and which templates embed them.
  const required = new Map<string, { embedded: boolean }>();
  for (const template of templates) {
    const pptxPath = path.join(TEMPLATE_ROOT, template, "template.pptx");
    if (!(await exists(pptxPath))) continue;
    const pkg = await PptxPackage.load(pptxPath);
    const typefaces = await extractFonts(pkg);
    const embedded = await embeddedFonts(pkg);
    for (const typeface of typefaces) {
      if (GENERIC_FONTS.has(typeface)) continue;
      const entry = required.get(typeface) ?? { embedded: false };
      if (embedded.has(typeface)) entry.embedded = true;
      required.set(typeface, entry);
    }
  }

  // 2. Check the system font registry.
  const installed = await listInstalledFonts();
  const statuses: FontStatus[] = [...required.entries()]
    .map(([typeface, info]) => ({
      typeface,
      family: baseFamily(typeface),
      embedded: info.embedded,
      system: isOnSystem(typeface, installed)
    }))
    .sort((a, b) => a.typeface.localeCompare(b.typeface));

  console.log(`Scanned ${templates.length} templates. ${statuses.length} distinct typefaces required.\n`);
  console.log("typeface                     embedded   system");
  console.log("------------------------------------------------");
  for (const status of statuses) {
    console.log(
      `${status.typeface.padEnd(28)} ${status.embedded ? "yes" : "no "}        ${status.system ? "yes" : "no "}`
    );
  }
  console.log("");

  // 3. Remove any stale variable brand files. LibreOffice substitutes these with
  //    a wrong font, and a leftover variable file can shadow the static one.
  await mkdir(USER_FONT_DIR, { recursive: true });
  const removedStale: string[] = [];
  for (const fileName of STALE_VARIABLE_FILES) {
    const target = path.join(USER_FONT_DIR, fileName);
    if (await exists(target)) {
      await rm(target, { force: true });
      removedStale.push(fileName);
    }
  }
  if (removedStale.length) {
    console.log(`Removed ${removedStale.length} stale variable font file(s): ${removedStale.join(", ")}`);
  }

  // 4. Install static instances of each brand family from the Google Fonts CSS2
  //    API. Static files carry clean family names that LibreOffice resolves.
  const installedNow: string[] = [];
  const cannotInstall: string[] = [];

  for (const [family, spec] of Object.entries(BRAND_FONTS)) {
    let faces: StaticFace[];
    try {
      faces = await fetchStaticFaces(family, spec.weights, spec.italicWeights);
    } catch (error) {
      console.warn(`Could not look up static files for ${family}: ${(error as Error).message}`);
      cannotInstall.push(family);
      continue;
    }
    if (faces.length === 0) {
      console.warn(`No static .ttf faces found for ${family} (the Google Fonts API response was empty or in an unexpected format).`);
      cannotInstall.push(family);
      continue;
    }
    for (const face of faces) {
      const fileName = staticFileName(face);
      const destination = path.join(USER_FONT_DIR, fileName);
      if (await exists(destination)) continue;
      try {
        const buffer = await download(face.url);
        await writeFile(destination, buffer);
        installedNow.push(fileName);
        console.log(`Installed ${fileName}`);
      } catch (error) {
        console.warn(`Could not download ${fileName}: ${(error as Error).message}`);
        cannotInstall.push(family);
      }
    }
  }

  // 5. Summarise.
  console.log("\nSummary");
  console.log("-------");
  const notCoveredAnywhere = statuses.filter((status) => !status.embedded && !status.system);
  if (installedNow.length) console.log(`Installed ${installedNow.length} static font file(s) this run.`);

  if (notCoveredAnywhere.length === 0) {
    console.log("Every required font is either embedded in its template or installed on this machine.");
    console.log("Decks will render correctly: missing system fonts are carried inside the output PPTX.");
  } else {
    console.log("These fonts are NOT embedded and NOT installed (hard gap):");
    for (const status of notCoveredAnywhere) console.log(`  - ${status.typeface}`);
    console.log("Options: install them manually, or accept the renderer's substitution fallback.");
  }

  const stillMissing = [...new Set(cannotInstall)].filter((family) => !BRAND_FONTS[family]);
  if (stillMissing.length) {
    console.log(`\nNo catalogue entry to auto-install: ${stillMissing.join(", ")}.`);
    console.log("Add the family to BRAND_FONTS in src/install-fonts.ts to support it.");
  }

  if (installedNow.length || removedStale.length) {
    console.log("\nNote: quit any running LibreOffice/soffice process so it picks up the new fonts.");
    console.log("macOS may also need a moment (or an app restart) to register newly installed fonts.");
  }
}

// Ask the Google Fonts CSS2 API for the static .ttf URL of each requested weight
// (upright and italic). Returns one entry per available face.
async function fetchStaticFaces(family: string, weights: number[], italicWeights: number[]): Promise<StaticFace[]> {
  const tuples: string[] = [];
  for (const w of weights) tuples.push(`0,${w}`);
  for (const w of italicWeights) tuples.push(`1,${w}`);
  tuples.sort();
  const familyParam = `${family.replace(/ /g, "+")}:ital,wght@${tuples.join(";")}`;
  const url = `https://fonts.googleapis.com/css2?family=${familyParam}`;

  const response = await fetch(url, { headers: { "User-Agent": LEGACY_USER_AGENT } });
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  const css = await response.text();
  return parseCssFaces(css);
}

// Pull (family, style, weight, ttf url) out of each @font-face block. The
// font-style line is omitted for upright faces, so it defaults to normal.
function parseCssFaces(css: string): StaticFace[] {
  const faces: StaticFace[] = [];
  for (const block of css.split("@font-face").slice(1)) {
    const familyMatch = block.match(/font-family:\s*'([^']+)'/);
    const weightMatch = block.match(/font-weight:\s*(\d+)/);
    const urlMatch = block.match(/src:\s*url\(([^)]+\.ttf)\)/);
    if (!familyMatch || !weightMatch || !urlMatch) continue;
    const italic = /font-style:\s*italic/.test(block);
    faces.push({ family: familyMatch[1], weight: Number.parseInt(weightMatch[1], 10), italic, url: urlMatch[1] });
  }
  return faces;
}

function staticFileName(face: StaticFace): string {
  const weightWord = WEIGHT_NAMES[face.weight] ?? String(face.weight);
  return `${face.family.replace(/ /g, "")}-${weightWord}${face.italic ? "Italic" : ""}.ttf`;
}

async function embeddedFonts(pkg: PptxPackage): Promise<Set<string>> {
  const xml = await pkg.text("ppt/presentation.xml");
  const embedded = new Set<string>();
  for (const match of xml.matchAll(/<p:embeddedFont>[\s\S]*?<p:font typeface="([^"]+)"\/>[\s\S]*?<\/p:embeddedFont>/g)) {
    embedded.add(match[1]);
  }
  return embedded;
}

function isOnSystem(typeface: string, installed: string): boolean {
  const haystack = installed.toLowerCase();
  if (haystack.includes(typeface.toLowerCase())) return true;
  // A variable family (e.g. "Inter") covers all weight-named requests.
  return haystack.includes(baseFamily(typeface).toLowerCase());
}

async function listInstalledFonts(): Promise<string> {
  // Prefer fontconfig (fast, precise); fall back to the macOS profiler.
  try {
    const result = await execFileAsync("fc-list", [], { maxBuffer: 50 * 1024 * 1024 });
    if (result.stdout.trim()) return result.stdout;
  } catch {
    // fontconfig not present; fall through.
  }
  try {
    const result = await execFileAsync("system_profiler", ["SPFontsDataType"], { maxBuffer: 50 * 1024 * 1024 });
    return result.stdout;
  } catch {
    return "";
  }
}

async function download(url: string): Promise<Buffer> {
  const response = await fetch(url, { headers: { "User-Agent": LEGACY_USER_AGENT } });
  if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
  const buffer = Buffer.from(await response.arrayBuffer());
  if (buffer.length < 10_000) throw new Error(`file unexpectedly small (${buffer.length} bytes)`);
  return buffer;
}

async function exists(filePath: string): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}

await main();
