import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { access, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { pdf } from "pdf-to-img";
import type { BuildWarning } from "./types.js";

const execFileAsync = promisify(execFile);

// Screenshot rendering is an OPTIONAL quality-assurance step. Building the .pptx
// itself needs no external binaries. Screenshots need a PPTX renderer, and the
// one reliable, free, cross-platform option is LibreOffice: it converts the deck
// to a PDF, and `pdf-to-img` (pure JS, no Poppler) turns each page into a PNG.
//
// If LibreOffice is not installed, we skip screenshots with a clear message
// instead of failing. Install it from https://www.libreoffice.org to enable
// them, then re-run the build.
const SOFFICE_CANDIDATES = [
  "soffice",
  "libreoffice",
  "/Applications/LibreOffice.app/Contents/MacOS/soffice"
];

export async function renderScreenshots(pptxPath: string, outputDir: string, warnings: BuildWarning[]): Promise<string[]> {
  await mkdir(outputDir, { recursive: true });
  await removeExistingScreenshots(outputDir);

  const soffice = await findSoffice();
  if (!soffice) {
    warnings.push({
      code: "screenshots-skipped",
      message: "LibreOffice was not found, so screenshots were skipped. The .pptx was still built. Install LibreOffice (https://www.libreoffice.org) to enable screenshots."
    });
    return [];
  }

  try {
    await execFileAsync(soffice, ["--headless", "--convert-to", "pdf", "--outdir", outputDir, pptxPath]);
  } catch (error) {
    warnings.push({
      code: "screenshots-skipped",
      message: `LibreOffice failed to convert the deck to PDF: ${errorMessage(error)}`
    });
    return [];
  }

  const pdfPath = path.join(outputDir, `${path.basename(pptxPath, ".pptx")}.pdf`);
  try {
    return await renderPdfToPngs(pdfPath, outputDir);
  } catch (error) {
    warnings.push({
      code: "screenshots-skipped",
      message: `PDF-to-PNG rendering failed: ${errorMessage(error)}`
    });
    return [];
  }
}

async function renderPdfToPngs(pdfPath: string, outputDir: string): Promise<string[]> {
  const document = await pdf(pdfPath, { scale: 2 });
  const screenshots: string[] = [];

  try {
    let pageNumber = 1;
    for await (const image of document) {
      const outputPath = path.join(outputDir, `slide-${String(pageNumber).padStart(2, "0")}.png`);
      await writeFile(outputPath, image);
      screenshots.push(outputPath);
      pageNumber += 1;
    }
  } finally {
    document.destroy();
  }

  if (screenshots.length === 0) {
    throw new Error("pdf-to-img produced no PNG pages.");
  }

  return screenshots;
}

async function findSoffice(): Promise<string | undefined> {
  for (const candidate of SOFFICE_CANDIDATES) {
    if (candidate.includes("/")) {
      try {
        await access(candidate);
        return candidate;
      } catch {
        continue;
      }
    }
    try {
      const result = await execFileAsync(process.platform === "win32" ? "where" : "which", [candidate]);
      const commandPath = result.stdout.trim().split(/\r?\n/)[0];
      if (commandPath) return commandPath;
    } catch {
      continue;
    }
  }
  return undefined;
}

async function removeExistingScreenshots(outputDir: string): Promise<void> {
  const files = await readdir(outputDir).catch(() => []);
  await Promise.all(
    files
      .filter((file) => file.endsWith(".png") || file.endsWith(".pdf"))
      .map((file) => rm(path.join(outputDir, file), { force: true }))
  );
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
