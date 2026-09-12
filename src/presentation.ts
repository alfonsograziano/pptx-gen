import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import type { AddSlideOptions, BuildReport, BuildReportSlide, BuildWarning, DeckSlide, RenderOptions } from "./types.js";
import { PptxPackage } from "./pptx-package.js";
import { appendSlideFromPackage, applyOverrides, applySlideNumbering, convertCustomSlidePageNumber, convertTemplatePageNumber, fillSlideText, flattenSlideNumberFields, getSlideEntries, keepOnlySlides, mergeEmbeddedFonts, validateFonts, validatePackage } from "./ooxml.js";
import { richTextToPlain } from "./rich-text.js";
import { loadTemplate } from "./templates.js";
import { ensureDir, writeTextFile } from "./fs.js";
import { renderScreenshots } from "./render.js";
import { CustomSlide, makeCustomSlideTempPath, renderCustomSlideToPptx, renderCustomSlidesToPptx } from "./custom-slide.js";
import { BuildProgress, dimKind } from "./progress.js";

// A no-op reporter used when progress output is disabled, so the render path
// can call the same methods without branching everywhere.
const SILENT_PROGRESS = {
  step: async <T>(_label: string, fn: () => Promise<T> | T) => fn(),
  note: (_message: string) => {},
  finish: (_summary: { output: string; slides: number; warnings: number }) => {}
};

export type PresentationOptions = {
  title?: string;
  templateLibrary?: string;
  projectDir?: string;
  /** Folder holding icons and logo assets. Defaults to `<templateLibrary>/../assets`. */
  assetsDir?: string;
};

export class Presentation {
  private readonly slides: DeckSlide[] = [];
  private readonly templateRoot: string;
  private readonly projectDir: string;
  private readonly title?: string;
  private readonly assetsDir: string;

  constructor(options: PresentationOptions = {}) {
    this.title = options.title;
    this.templateRoot = path.resolve(options.templateLibrary ?? "templates");
    this.projectDir = path.resolve(options.projectDir ?? process.cwd());
    this.assetsDir = path.resolve(options.assetsDir ?? path.resolve(this.templateRoot, "..", "assets"));
  }

  addSlideFromTemplate(options: AddSlideOptions): this {
    this.slides.push({ kind: "template", options });
    return this;
  }

  addCustomSlide(slide: CustomSlide): this {
    this.slides.push({ kind: "custom", slide });
    return this;
  }

  async render(options: RenderOptions): Promise<BuildReport> {
    if (this.slides.length === 0) throw new Error("Cannot render a presentation with no slides.");

    const output = path.resolve(this.projectDir, options.output);
    const warnings: BuildWarning[] = [];
    const listing = buildSlideListing(this.slides, warnings);
    const firstTemplateSlide = this.slides.find((slide) => slide.kind === "template");

    if (!firstTemplateSlide) {
      return this.renderAllCustom(output, options, warnings, listing);
    }

    // Steps: one per slide + validate fonts + save + validate package + screenshots (+ report).
    const expectedSteps = this.slides.length + 4 + (options.report ? 1 : 0);
    const progress = options.progress === false
      ? SILENT_PROGRESS
      : new BuildProgress(this.title ?? "Building deck", expectedSteps);

    const firstTemplate = await loadTemplate(this.templateRoot, firstTemplateSlide.options.templateName);
    const pkg = await PptxPackage.load(firstTemplate.pptxPath);
    // The base package doubles as the source for slides reusing the first
    // template, and it grows as we append. Capture its slides now so a later
    // reuse still sees the template's own slides rather than the whole deck.
    const firstTemplateEntries = await getSlideEntries(pkg);
    const clonedSlides: number[] = [];
    const templatesUsed: string[] = [];
    const customSlidesUsed: string[] = [];
    const requiredFonts = new Set<string>();

    // Each template package is a single-slide slice that shares an identical
    // support chain (layouts, masters, themes, fonts) with the base package.
    // We merge every requested slide into the base package as a new slide, then
    // trim the base down to only the slides we built.
    for (const [index, requestedSlide] of this.slides.entries()) {
      const position = `${index + 1}/${this.slides.length}`;

      if (requestedSlide.kind === "custom") {
        await progress.step(`Slide ${position}  ${requestedSlide.slide.name} ${dimKind(kindLabel(listing[index]))}`, async () => {
          for (const font of requestedSlide.slide.requiredFonts) requiredFonts.add(font);

          const tempPptx = await makeCustomSlideTempPath();
          await renderCustomSlideToPptx({
            customSlide: requestedSlide.slide,
            output: tempPptx,
            pageNum: index + 1,
            projectDir: this.projectDir,
            assetsDir: this.assetsDir,
            title: this.title
          });

          const srcPkg = await PptxPackage.load(tempPptx);
          const srcEntries = await getSlideEntries(srcPkg);
          if (srcEntries.length === 0) {
            throw new Error(`Custom slide '${requestedSlide.slide.name}' produced no slides.`);
          }

          const clonedSlideNumber = await appendSlideFromPackage(pkg, srcPkg, srcEntries[0].slideNumber, warnings);
          await convertCustomSlidePageNumber(pkg, clonedSlideNumber);
          clonedSlides.push(clonedSlideNumber);
          customSlidesUsed.push(requestedSlide.slide.name);
          warnings.push({
            code: "custom-slide-generated",
            message: `Generated custom slide '${requestedSlide.slide.name}'.`,
            slide: index + 1,
            target: requestedSlide.slide.name
          });
        });
        continue;
      }

      await progress.step(`Slide ${position}  ${requestedSlide.options.templateName} ${dimKind(kindLabel(listing[index]))}`, async () => {
        const template = requestedSlide.options.templateName === firstTemplate.id
          ? firstTemplate
          : await loadTemplate(this.templateRoot, requestedSlide.options.templateName);
        for (const font of template.metadata.fonts ?? []) requiredFonts.add(font);

        const isFirstTemplate = template.id === firstTemplate.id;
        const srcPkg = isFirstTemplate ? pkg : await PptxPackage.load(template.pptxPath);
        const srcEntries = isFirstTemplate ? firstTemplateEntries : await getSlideEntries(srcPkg);
        if (srcEntries.length === 0) {
          throw new Error(`Slide ${index + 1} template '${requestedSlide.options.templateName}' contains no slides.`);
        }
        if (srcEntries.length > 1) {
          warnings.push({
            code: "multi-slide-template",
            message: `Template '${requestedSlide.options.templateName}' contains ${srcEntries.length} slides; using the first.`,
            slide: index + 1,
            target: requestedSlide.options.templateName
          });
        }

        const clonedSlideNumber = await appendSlideFromPackage(pkg, srcPkg, srcEntries[0].slideNumber, warnings);
        // Carry over any fonts this slide's source embeds that the base package
        // lacks, so the output deck is self-contained for every typeface it uses.
        await mergeEmbeddedFonts(pkg, srcPkg, warnings);
        clonedSlides.push(clonedSlideNumber);
        templatesUsed.push(requestedSlide.options.templateName);

        const variables = Object.fromEntries(
          Object.entries(requestedSlide.options.variables ?? {}).map(([key, value]) => [key, richTextToPlain(value)])
        );

        await fillSlideText(pkg, clonedSlideNumber, template.fieldsFile.fields, variables, warnings);
        await convertTemplatePageNumber(pkg, clonedSlideNumber, template.fieldsFile.fields);
        await applyOverrides(pkg, clonedSlideNumber, template.fieldsFile.fields, requestedSlide.options.overrides ?? [], this.projectDir, warnings);
      });
    }

    await progress.step(`Validating fonts ${dimKind([...requiredFonts].join(", ") || "none")}`, () => validateFonts(pkg, [...requiredFonts], warnings));
    await progress.step("Assembling and saving deck", async () => {
      await keepOnlySlides(pkg, clonedSlides);
      // Numbering can only be settled once the slides are in their final order.
      await applySlideNumbering(pkg, warnings);
      await ensureDir(path.dirname(output));
      await pkg.save(output);
    });
    await progress.step("Validating PPTX package", () => validatePackage(output));

    const screenshotDir = options.screenshots
      ? path.resolve(this.projectDir, options.screenshots)
      : path.join(path.dirname(output), "screenshots");
    const screenshots = await progress.step(
      "Rendering screenshots (LibreOffice)",
      () => this.renderPreview(output, screenshotDir, warnings)
    );

    const report: BuildReport = {
      generatedAt: new Date().toISOString(),
      output,
      templatesUsed,
      customSlidesUsed,
      slides: listing,
      slidesBuilt: this.slides.length,
      warnings,
      screenshots
    };

    if (options.report) {
      await progress.step("Writing build report", () =>
        writeTextFile(path.resolve(this.projectDir, options.report!), formatReport(report))
      );
    }

    progress.finish({ output, slides: report.slidesBuilt, warnings: warnings.length });

    return report;
  }

  /**
   * Screenshot a throwaway copy of the deck with its slide-number fields
   * flattened to plain text.
   *
   * LibreOffice resolves `slidenum` fields but ignores the deck's
   * `firstSlideNum` offset, so shooting the delivered file would show a number
   * one higher than PowerPoint on every slide of a deck with a cover. The
   * delivered .pptx keeps its live fields; only this copy is flattened.
   */
  private async renderPreview(output: string, screenshotDir: string, warnings: BuildWarning[]): Promise<string[]> {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "pptx-gen-preview-"));
    try {
      const flattened = path.join(tempDir, path.basename(output));
      const pkg = await PptxPackage.load(output);
      await flattenSlideNumberFields(pkg);
      await pkg.save(flattened);
      return await renderScreenshots(flattened, screenshotDir, warnings);
    } catch {
      // Previewing is a convenience, so fall back to the delivered file rather
      // than losing screenshots over it.
      return renderScreenshots(output, screenshotDir, warnings);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  }

  private async renderAllCustom(
    output: string,
    options: RenderOptions,
    warnings: BuildWarning[],
    listing: BuildReportSlide[]
  ): Promise<BuildReport> {
    const customSlides = this.slides.map((slide) => {
      if (slide.kind !== "custom") throw new Error("Unexpected non-custom slide in all-custom render path.");
      return slide.slide;
    });

    // Steps: render+save + validate package + screenshots (+ report).
    const expectedSteps = 3 + (options.report ? 1 : 0);
    const progress = options.progress === false
      ? SILENT_PROGRESS
      : new BuildProgress(this.title ?? "Building deck", expectedSteps);

    await progress.step(`Rendering ${customSlides.length} custom slide(s) and saving deck`, async () => {
      await renderCustomSlidesToPptx({
        customSlides,
        output,
        projectDir: this.projectDir,
        assetsDir: this.assetsDir,
        title: this.title
      });

      // pptxgenjs writes the file directly, so the page-number text boxes become
      // live fields in a second pass over the saved package.
      const pkg = await PptxPackage.load(output);
      for (const entry of await getSlideEntries(pkg)) {
        await convertCustomSlidePageNumber(pkg, entry.slideNumber);
      }
      await applySlideNumbering(pkg, warnings);
      await pkg.save(output);
    });
    await progress.step("Validating PPTX package", () => validatePackage(output));

    customSlides.forEach((slide, index) => {
      warnings.push({
        code: "custom-slide-generated",
        message: `Generated custom slide '${slide.name}'.`,
        slide: index + 1,
        target: slide.name
      });
    });

    const screenshotDir = options.screenshots
      ? path.resolve(this.projectDir, options.screenshots)
      : path.join(path.dirname(output), "screenshots");
    const screenshots = await progress.step(
      "Rendering screenshots (LibreOffice)",
      () => this.renderPreview(output, screenshotDir, warnings)
    );

    const report: BuildReport = {
      generatedAt: new Date().toISOString(),
      output,
      templatesUsed: [],
      customSlidesUsed: customSlides.map((slide) => slide.name),
      slides: listing,
      slidesBuilt: customSlides.length,
      warnings,
      screenshots
    };

    if (options.report) {
      await progress.step("Writing build report", () =>
        writeTextFile(path.resolve(this.projectDir, options.report!), formatReport(report))
      );
    }

    progress.finish({ output, slides: report.slidesBuilt, warnings: warnings.length });

    return report;
  }
}

/**
 * Describe every slide in output order, resolving variant-group membership.
 *
 * Slides tagged with the same `group` are alternative takes on one concept, so
 * the reviewer can compare them and pick one. A group of one is not a variant
 * set, so its tag is dropped rather than reported as "variant 1 of 1".
 */
function buildSlideListing(slides: DeckSlide[], warnings: BuildWarning[]): BuildReportSlide[] {
  const listing: BuildReportSlide[] = slides.map((slide, index) =>
    slide.kind === "custom"
      ? { index: index + 1, kind: "custom", name: slide.slide.name, group: slide.slide.group }
      : { index: index + 1, kind: "template", name: slide.options.templateName, group: slide.options.group }
  );

  const positionsByGroup = new Map<string, number[]>();
  for (const entry of listing) {
    if (!entry.group) continue;
    const positions = positionsByGroup.get(entry.group) ?? [];
    positions.push(entry.index);
    positionsByGroup.set(entry.group, positions);
  }

  const seen = new Map<string, number>();
  for (const entry of listing) {
    if (!entry.group) continue;
    const positions = positionsByGroup.get(entry.group)!;
    if (positions.length < 2) {
      delete entry.group;
      continue;
    }
    const variant = (seen.get(entry.group) ?? 0) + 1;
    seen.set(entry.group, variant);
    entry.variant = variant;
    entry.variantCount = positions.length;
  }

  // Variants only read as alternatives when they sit next to each other; a split
  // group almost always means the build script added an unrelated slide between
  // them. Warn rather than fail, matching the other build-time warnings.
  for (const [group, positions] of positionsByGroup) {
    if (positions.length < 2) continue;
    const contiguous = positions[positions.length - 1] - positions[0] === positions.length - 1;
    if (contiguous) continue;
    warnings.push({
      code: "variant-group-split",
      message: `Variant group '${group}' is not consecutive; its slides are at positions ${positions.join(", ")}.`,
      target: group
    });
  }

  return listing;
}

/** Dim annotation for a slide's progress line, e.g. "custom · agenda 2/3". */
function kindLabel(entry: BuildReportSlide): string {
  return entry.group ? `${entry.kind} · ${entry.group} ${entry.variant}/${entry.variantCount}` : entry.kind;
}

function formatSlidesSection(report: BuildReport): string {
  // Only pair slides with screenshots when there is one per slide; LibreOffice
  // may have been missing, in which case the suffix is simply left off.
  const screenshots = report.screenshots.length === report.slides.length ? report.screenshots : null;
  const shot = (entry: BuildReportSlide) =>
    screenshots ? ` — ${path.basename(screenshots[entry.index - 1])}` : "";

  const lines: string[] = [];
  for (let i = 0; i < report.slides.length; i += 1) {
    const entry = report.slides[i];
    if (!entry.group) {
      lines.push(`- ${entry.index}. ${entry.name} (${entry.kind})${shot(entry)}`);
      continue;
    }
    if (entry.variant === 1) {
      const last = entry.index + entry.variantCount! - 1;
      lines.push(
        `- Variant group \`${entry.group}\` — ${entry.variantCount} variants, slides ${entry.index}-${last}. Pick one:`
      );
    }
    lines.push(
      `  - ${entry.index}. ${entry.name} (${entry.kind}) — variant ${entry.variant} of ${entry.variantCount}${shot(entry)}`
    );
  }
  return lines.join("\n");
}

function formatVariantGroups(report: BuildReport): string {
  const counts = new Map<string, number>();
  for (const entry of report.slides) {
    if (!entry.group) continue;
    counts.set(entry.group, entry.variantCount!);
  }
  if (counts.size === 0) return "none";
  return [...counts].map(([group, count]) => `${group} (${count})`).join(", ");
}

function formatReport(report: BuildReport): string {
  return `# Build report

- Generated at: ${report.generatedAt}
- Output: ${report.output}
- Slides built: ${report.slidesBuilt}
- Templates used: ${report.templatesUsed.join(", ")}
- Custom slides used: ${report.customSlidesUsed.length ? report.customSlidesUsed.join(", ") : "none"}
- Variant groups: ${formatVariantGroups(report)}
- Screenshots: ${report.screenshots.length ? report.screenshots.join(", ") : "none"}

## Slides

${formatSlidesSection(report)}

## Warnings

${report.warnings.length ? report.warnings.map((warning) => `- ${warning.code}: ${warning.message}`).join("\n") : "- None"}
`;
}
