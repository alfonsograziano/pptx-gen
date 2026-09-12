import path from "node:path";
import { ensureDir, writeTextFile, writeYamlFile } from "./fs.js";
import { PptxPackage } from "./pptx-package.js";
import { extractFonts, extractTextFields, getSlideEntries, sliceToSingleSlide } from "./ooxml.js";
import type { FieldsFile, TemplateMetadata } from "./types.js";
import { renderScreenshots } from "./render.js";
import { resolveWorkspaceSync } from "./workspace.js";

export type IngestOptions = {
  source: string;
  templateName: string;
  slide?: number;
  split?: boolean;
  templateRoot?: string;
};

export async function ingestTemplate(options: IngestOptions): Promise<string[]> {
  const templateRoot = path.resolve(options.templateRoot ?? resolveWorkspaceSync().templatesDir);
  const source = path.resolve(options.source);
  const sourcePkg = await PptxPackage.load(source);
  const slides = await getSlideEntries(sourcePkg);

  if (slides.length > 1 && !options.split && options.slide === undefined) {
    throw new Error(
      `Source PPTX has ${slides.length} slides. Pass --split to import all slides as separate templates, or pass --slide <n> to import one slide.`
    );
  }

  if (options.split) {
    const imported: string[] = [];
    for (let slide = 1; slide <= slides.length; slide += 1) {
      const templateName = `${options.templateName}-slide-${String(slide).padStart(2, "0")}`;
      imported.push(await ingestOneSlide({ source, templateRoot, templateName, slide, totalSlides: slides.length }));
    }
    return imported;
  }

  const slide = options.slide ?? 1;
  return [await ingestOneSlide({ source, templateRoot, templateName: options.templateName, slide, totalSlides: slides.length })];
}

type IngestOneSlideOptions = {
  source: string;
  templateRoot: string;
  templateName: string;
  slide: number;
  totalSlides: number;
};

async function ingestOneSlide(options: IngestOneSlideOptions): Promise<string> {
  const templateDir = path.join(options.templateRoot, options.templateName);
  const templatePptx = path.join(templateDir, "template.pptx");
  await ensureDir(templateDir);

  const pkg = await PptxPackage.load(options.source);
  const keptSlideNumber = await sliceToSingleSlide(pkg, options.slide);
  await pkg.save(templatePptx);

  const singleSlidePkg = await PptxPackage.load(templatePptx);
  const fields = await extractTextFields(singleSlidePkg, keptSlideNumber);
  const fonts = await extractFonts(singleSlidePkg, keptSlideNumber);
  const pageNumberField = fields.find((field) => field.role === "page-number");

  const metadata: TemplateMetadata = {
    id: options.templateName,
    name: humanName(options.templateName),
    kind: "cloned-slide",
    status: "draft",
    version: "1.0.0",
    source: {
      // Relative to the template library, not the cwd, so the recorded path
      // means the same thing whatever directory the ingest was run from.
      deck: path.relative(options.templateRoot, options.source),
      slide: options.slide,
      importedOn: new Date().toISOString().slice(0, 10)
    },
    fonts,
    // The page number is the deck's to manage, not the author's, so it is not
    // offered as a variable.
    variables: fields
      .filter((field) => field.originalText.trim() && field.role !== "page-number")
      .map((field) => field.id),
    tags: []
  };

  const fieldsFile: FieldsFile = {
    templateId: options.templateName,
    librarySlide: 0,
    fields
  };

  await writeYamlFile(path.join(templateDir, "template.yml"), metadata);
  await writeYamlFile(path.join(templateDir, "fields.yml"), fieldsFile);
  await writeTextFile(path.join(templateDir, "description.md"), descriptionStub(options.templateName));
  await writeTextFile(path.join(templateDir, "ingestion-report.md"), ingestionReport({
    source: options.source,
    slide: options.slide,
    totalSlides: options.totalSlides,
    fonts,
    fieldCount: fields.length,
    pageNumberField: pageNumberField?.name
  }));

  const warnings: { code: string; message: string }[] = [];
  await renderScreenshots(templatePptx, path.join(templateDir, "screenshots"), warnings);
  if (warnings.length > 0) {
    await writeTextFile(path.join(templateDir, "screenshot-warnings.md"), warnings.map((warning) => `- ${warning.code}: ${warning.message}`).join("\n"));
  }

  return options.templateName;
}

function humanName(value: string): string {
  return value.replace(/[-_]+/g, " ").replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function descriptionStub(templateName: string): string {
  return `# ${humanName(templateName)}

## What this slide communicates

TODO.

## When to use it

TODO.

## When not to use it

TODO.

## Layout constraints

TODO.
`;
}

function ingestionReport(input: {
  source: string;
  slide: number;
  totalSlides: number;
  fonts: string[];
  fieldCount: number;
  pageNumberField?: string;
}): string {
  return `# Ingestion report

- Source: ${input.source}
- Source slide: ${input.slide}
- Source slide count: ${input.totalSlides}
- Text fields: ${input.fieldCount}
- Fonts: ${input.fonts.length ? input.fonts.join(", ") : "none detected"}
- Page number: ${input.pageNumberField
    ? `detected on shape '${input.pageNumberField}', tagged \`role: page-number\` in fields.yml and rendered as a live slide-number field`
    : "none detected — if this slide does show one, tag its field with `role: page-number` in fields.yml or it will ship frozen"}
- Status: imported
`;
}
