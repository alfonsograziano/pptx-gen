import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";
import { extractFonts, extractTextFields, getSlideEntries } from "./ooxml.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TITLE_TEMPLATE = path.resolve(HERE, "..", "templates", "title-cover", "template.pptx");

test("getSlideEntries finds the single slide in an ingested template", async () => {
  const pkg = await PptxPackage.load(TITLE_TEMPLATE);
  const entries = await getSlideEntries(pkg);
  assert.equal(entries.length, 1);
  assert.match(entries[0].target, /slides\/slide\d+\.xml$/);
});

test("extractTextFields returns stable, unique field ids", async () => {
  const pkg = await PptxPackage.load(TITLE_TEMPLATE);
  const entries = await getSlideEntries(pkg);
  const fields = await extractTextFields(pkg, entries[0].slideNumber);

  const ids = fields.map((f) => f.id);
  assert.ok(ids.includes("your-presentation-title-goes-here"), `ids were: ${ids.join(", ")}`);
  assert.equal(new Set(ids).size, ids.length, "field ids must be unique");
  for (const field of fields) {
    assert.equal(typeof field.shapeId, "string");
    assert.notEqual(field.shapeId, "");
  }
});

test("extractFonts lists real typefaces and drops theme references", async () => {
  const pkg = await PptxPackage.load(TITLE_TEMPLATE);
  const fonts = await extractFonts(pkg, (await getSlideEntries(pkg))[0].slideNumber);
  assert.ok(fonts.includes("Inter"));
  assert.ok(!fonts.some((f) => f.startsWith("+")), `theme refs leaked: ${fonts.join(", ")}`);
});
