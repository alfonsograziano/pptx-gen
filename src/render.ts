import { access, mkdir, readdir, rm, writeFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { pdf } from "pdf-to-img";
import type { BuildWarning } from "./types.js";

const execFileAsync = promisify(execFile);

export async function renderScreenshots(pptxPath: string, outputDir: string, warnings: BuildWarning[]): Promise<string[]> {
  await mkdir(outputDir, { recursive: true });
  await removeExistingScreenshots(outputDir);

  const soffice = await findExecutable([
    "soffice",
    "libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice"
  ]);

  if (!soffice) {
    warnings.push({
      code: "screenshots-skipped",
      message: "LibreOffice was not available, so screenshot rendering was skipped."
    });
    return [];
  }

  try {
    await execFileAsync(soffice, ["--headless", "--convert-to", "pdf", "--outdir", outputDir, pptxPath]);
  } catch (error) {
    warnings.push({
      code: "screenshots-skipped",
      message: `LibreOffice failed to convert PPTX to PDF: ${error instanceof Error ? error.message : String(error)}`
    });
    return [];
  }

  const pdfPath = path.join(outputDir, `${path.basename(pptxPath, ".pptx")}.pdf`);
  const pdftoppm = await findExecutable(["pdftoppm"]);

  if (pdftoppm) {
    try {
      await execFileAsync(pdftoppm, ["-png", "-r", "150", pdfPath, path.join(outputDir, "slide")]);
      return listPngs(outputDir);
    } catch {
      warnings.push({
        code: "pdftoppm-failed",
        message: "Poppler 'pdftoppm' failed, falling back to pdf-to-img."
      });
    }
  }

  try {
    return await renderPdfWithPdfToImg(pdfPath, outputDir);
  } catch (error) {
    warnings.push({
      code: "screenshots-skipped",
      message: `PDF-to-PNG rendering failed: ${error instanceof Error ? error.message : String(error)}`
    });
    return [];
  }
}

async function renderPdfWithPdfToImg(pdfPath: string, outputDir: string): Promise<string[]> {
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

async function findExecutable(candidates: string[]): Promise<string | undefined> {
  for (const candidate of candidates) {
    if (candidate.includes("/")) {
      try {
        await access(candidate);
        return candidate;
      } catch {
        continue;
      }
    }

    try {
      const result = await execFileAsync("zsh", ["-lc", `command -v ${candidate}`]);
      const commandPath = result.stdout.trim();
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

async function listPngs(outputDir: string): Promise<string[]> {
  const files = await readdir(outputDir);
  return files
    .filter((file) => file.endsWith(".png"))
    .sort()
    .map((file) => path.join(outputDir, file));
}
