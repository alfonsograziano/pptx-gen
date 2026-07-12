/**
 * End-to-end self-check: ingest a slide from the bundled example source deck,
 * build a deck from it with a replaced field and an addText override, and verify
 * the output is a valid single-slide package containing the new text. Runs fully
 * in a temp directory and needs no external binaries.
 *
 *   npm run self-validate
 */
import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { ingestTemplate, Presentation } from "./index.js";
import { readYamlFile } from "./fs.js";
import type { FieldsFile } from "./types.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const source = path.join(projectRoot, "examples", "example-source.pptx");
const templateName = "self-validation-title";

const workDir = await mkdtemp(path.join(os.tmpdir(), "pptx-gen-self-validate-"));
const templateRoot = path.join(workDir, "templates");
const deckProject = path.join(workDir, "deck");

try {
  // 1. Ingest one slide from the example source into a fresh template library.
  await ingestTemplate({ source, templateName, slide: 1, templateRoot });

  const fields = await readYamlFile<FieldsFile>(path.join(templateRoot, templateName, "fields.yml"));
  const titleField = fields.fields.find((field) => field.originalText.includes("presentation title"));
  if (!titleField) throw new Error("Self-validation fixture did not expose the expected title field.");

  // 2. Build a deck: replace the title and add a stamped text box.
  const deck = new Presentation({ templateLibrary: templateRoot, projectDir: deckProject });
  deck.addSlideFromTemplate({
    templateName,
    variables: { [titleField.id]: "Validated generation" },
    overrides: [
      {
        op: "addText",
        id: "validation-stamp",
        text: "Self-validation passed",
        x: 6.8,
        y: 0.35,
        w: 2.2,
        h: 0.25,
        style: { fontSize: 8, color: "5B6472", italic: true }
      }
    ]
  });

  const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

  // 3. Verify the output.
  const outputPath = path.join(deckProject, "deck.pptx");
  const pkg = await PptxPackage.load(outputPath);
  const entries = await getSlideEntries(pkg);
  if (entries.length !== 1) throw new Error(`Expected 1 output slide, got ${entries.length}.`);

  const slideXml = await pkg.text(`ppt/slides/slide${entries[0].slideNumber}.xml`);
  if (!slideXml.includes("Validated generation")) throw new Error("Output is missing the replaced title text.");
  if (!slideXml.includes("Self-validation passed")) throw new Error("Output is missing the addText override.");

  console.log("Self-validation passed.");
  console.log(`  Output: ${outputPath}`);
  console.log(`  Warnings: ${report.warnings.length ? report.warnings.map((w) => w.code).join(", ") : "none"}`);
} finally {
  await rm(workDir, { recursive: true, force: true });
}
