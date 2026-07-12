import path from "node:path";
import type { AddSlideOptions, BuildReport, BuildWarning, DeckSlide, RenderOptions } from "./types.js";
import { PptxPackage } from "./pptx-package.js";
import { appendSlideFromPackage, applyOverrides, fillSlideText, getSlideEntries, keepOnlySlides, mergeEmbeddedFonts, validateFonts, validatePackage } from "./ooxml.js";
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
    const firstTemplateSlide = this.slides.find((slide) => slide.kind === "template");

    if (!firstTemplateSlide) {
      return this.renderAllCustom(output, options, warnings);
    }

    // Steps: one per slide + validate fonts + save + validate package + screenshots (+ report).
    const expectedSteps = this.slides.length + 4 + (options.report ? 1 : 0);
    const progress = options.progress === false
      ? SILENT_PROGRESS
      : new BuildProgress(this.title ?? "Building deck", expectedSteps);

    const firstTemplate = await loadTemplate(this.templateRoot, firstTemplateSlide.options.templateName);
    const pkg = await PptxPackage.load(firstTemplate.pptxPath);
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
        await progress.step(`Slide ${position}  ${requestedSlide.slide.name} ${dimKind("custom")}`, async () => {
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

      await progress.step(`Slide ${position}  ${requestedSlide.options.templateName} ${dimKind("template")}`, async () => {
        const template = requestedSlide.options.templateName === firstTemplate.id
          ? firstTemplate
          : await loadTemplate(this.templateRoot, requestedSlide.options.templateName);
        for (const font of template.metadata.fonts ?? []) requiredFonts.add(font);

        const srcPkg = template.id === firstTemplate.id ? pkg : await PptxPackage.load(template.pptxPath);
        const srcEntries = await getSlideEntries(srcPkg);
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
        await applyOverrides(pkg, clonedSlideNumber, template.fieldsFile.fields, requestedSlide.options.overrides ?? [], this.projectDir, warnings);
      });
    }

    await progress.step(`Validating fonts ${dimKind([...requiredFonts].join(", ") || "none")}`, () => validateFonts(pkg, [...requiredFonts]));
    await progress.step("Assembling and saving deck", async () => {
      await keepOnlySlides(pkg, clonedSlides);
      await ensureDir(path.dirname(output));
      await pkg.save(output);
    });
    await progress.step("Validating PPTX package", () => validatePackage(output));

    const screenshotDir = options.screenshots
      ? path.resolve(this.projectDir, options.screenshots)
      : path.join(path.dirname(output), "screenshots");
    const screenshots = await progress.step(
      "Rendering screenshots (LibreOffice)",
      () => renderScreenshots(output, screenshotDir, warnings)
    );

    const report: BuildReport = {
      generatedAt: new Date().toISOString(),
      output,
      templatesUsed,
      customSlidesUsed,
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

  private async renderAllCustom(output: string, options: RenderOptions, warnings: BuildWarning[]): Promise<BuildReport> {
    const customSlides = this.slides.map((slide) => {
      if (slide.kind !== "custom") throw new Error("Unexpected non-custom slide in all-custom render path.");
      return slide.slide;
    });

    // Steps: render+save + validate package + screenshots (+ report).
    const expectedSteps = 3 + (options.report ? 1 : 0);
    const progress = options.progress === false
      ? SILENT_PROGRESS
      : new BuildProgress(this.title ?? "Building deck", expectedSteps);

    await progress.step(`Rendering ${customSlides.length} custom slide(s) and saving deck`, () =>
      renderCustomSlidesToPptx({
        customSlides,
        output,
        projectDir: this.projectDir,
        assetsDir: this.assetsDir,
        title: this.title
      })
    );
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
      () => renderScreenshots(output, screenshotDir, warnings)
    );

    const report: BuildReport = {
      generatedAt: new Date().toISOString(),
      output,
      templatesUsed: [],
      customSlidesUsed: customSlides.map((slide) => slide.name),
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

function formatReport(report: BuildReport): string {
  return `# Build report

- Generated at: ${report.generatedAt}
- Output: ${report.output}
- Slides built: ${report.slidesBuilt}
- Templates used: ${report.templatesUsed.join(", ")}
- Custom slides used: ${report.customSlidesUsed.length ? report.customSlidesUsed.join(", ") : "none"}
- Screenshots: ${report.screenshots.length ? report.screenshots.join(", ") : "none"}

## Warnings

${report.warnings.length ? report.warnings.map((warning) => `- ${warning.code}: ${warning.message}`).join("\n") : "- None"}
`;
}
