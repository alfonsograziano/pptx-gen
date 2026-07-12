/**
 * Re-extract a template's fields.yml from its existing template.pptx.
 *
 * Field extraction is deterministic, so text-field ids stay stable; this just
 * picks up anything a newer extractor now understands (for example image fields
 * for pictures, which are editable via the `replaceImage` override). Only
 * fields.yml is rewritten. template.yml, description.md, and screenshots are
 * left untouched.
 *
 * Usage: tsx src/reextract-fields.ts <template-name> [<template-name> ...]
 */
import path from "node:path";
import { readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";
import { extractTextFields, getSlideEntries } from "./ooxml.js";
import { writeYamlFile } from "./fs.js";
import type { FieldsFile } from "./types.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE_ROOT = path.resolve(HERE, "..", "templates");

// With no arguments, re-extract every template in the library; otherwise only
// the named ones.
const names = process.argv.slice(2).length > 0
  ? process.argv.slice(2)
  : (await readdir(TEMPLATE_ROOT, { withFileTypes: true }))
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name)
      .sort();

let updated = 0;
let totalImages = 0;
for (const name of names) {
  const dir = path.join(TEMPLATE_ROOT, name);
  const pptxPath = path.join(dir, "template.pptx");
  try {
    const pkg = await PptxPackage.load(pptxPath);
    const slides = await getSlideEntries(pkg);
    if (slides.length === 0) {
      console.error(`${name}: template.pptx has no slides, skipping.`);
      continue;
    }
    const fields = await extractTextFields(pkg, slides[0].slideNumber);
    const fieldsFile: FieldsFile = { templateId: name, librarySlide: 0, fields };
    await writeYamlFile(path.join(dir, "fields.yml"), fieldsFile);
    const images = fields.filter((field) => field.type === "image").length;
    totalImages += images;
    updated += 1;
    console.log(`${name}: ${fields.length} fields (${images} image, ${fields.length - images} text)`);
  } catch (error) {
    console.error(`${name}: failed to re-extract (${(error as Error).message})`);
  }
}
console.log(`\nDone. Updated ${updated}/${names.length} templates; ${totalImages} image field(s) total.`);
