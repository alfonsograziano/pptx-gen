import { mkdir, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { ingestTemplate, Presentation, md } from "./index.js";
import { readYamlFile } from "./fs.js";
import type { FieldsFile } from "./types.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(projectRoot, "..", "..");
const fixture = path.join(repoRoot, "lake", "Untitled presentation.pptx");
const workDir = path.join(projectRoot, "tmp", "self-validation");
const templateRoot = path.join(workDir, "templates");
const deckProject = path.join(workDir, "projects", "lake-import");
const templateName = "lake-four-step-funnel";

await rm(workDir, { recursive: true, force: true });
await mkdir(deckProject, { recursive: true });

await ingestTemplate({
  source: fixture,
  templateName,
  slide: 1,
  templateRoot
});

const fields = await readYamlFile<FieldsFile>(path.join(templateRoot, templateName, "fields.yml"));
const titleField = fields.fields.find((field) => field.originalText.includes("Progressive discovery"));
const introField = fields.fields.find((field) => field.originalText.includes("narrows the search space"));
if (!titleField || !introField) throw new Error("Self-validation fixture did not expose expected text fields.");

const deck = new Presentation({
  templateLibrary: templateRoot,
  projectDir: deckProject
});

deck.addSlideFromTemplate({
  templateName,
  variables: {
    [titleField.id]: "Validated generation_",
    [introField.id]: md("This slide was **ingested**, cloned, filled, and rebuilt by the TypeScript generator.")
  },
  overrides: [
    {
      op: "addText",
      id: "validation-stamp",
      text: "Self-validation passed",
      x: 6.8,
      y: 0.35,
      w: 2.2,
      h: 0.25,
      style: { fontFace: "Inter", fontSize: 8, color: "526288", italic: true }
    }
  ]
});

const report = await deck.render({
  output: "output/lake-import-generated.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots"
});

const outputPath = path.join(deckProject, "output", "lake-import-generated.pptx");
const pkg = await PptxPackage.load(outputPath);
const slideEntries = await getSlideEntries(pkg);
if (slideEntries.length !== 1) throw new Error(`Expected 1 output slide, got ${slideEntries.length}.`);
const outputText = await pkg.text(`ppt/slides/slide${slideEntries[0].slideNumber}.xml`);
if (!outputText.includes("Validated generation")) throw new Error("Generated PPTX does not include replaced title text.");
if (!outputText.includes("Self-validation passed")) throw new Error("Generated PPTX does not include addText override.");
const templateScreenshots = await listPngs(path.join(templateRoot, templateName, "screenshots"));
if (templateScreenshots.length < 1) throw new Error("Ingestion did not generate template PNG screenshots.");
const projectScreenshots = await listPngs(path.join(deckProject, "output", "screenshots"));
if (projectScreenshots.length < 1) throw new Error("Render did not generate project PNG screenshots.");

await writeFile(path.join(workDir, "SELF_VALIDATION_PASSED.txt"), [
  "Self-validation passed.",
  `Output: ${outputPath}`,
  `Template screenshots: ${templateScreenshots.length}`,
  `Project screenshots: ${projectScreenshots.length}`,
  `Warnings: ${report.warnings.length}`
].join("\n"), "utf8");

console.log(`Self-validation passed: ${outputPath}`);

async function listPngs(dir: string): Promise<string[]> {
  const files = await readdir(dir).catch(() => []);
  return files.filter((file) => file.endsWith(".png")).sort();
}
