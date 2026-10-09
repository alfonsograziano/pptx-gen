import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";
import { extractFonts, extractTextFields, getSlideEntries, roundStroke } from "./ooxml.js";
import { STARTER_TEMPLATES } from "./test-fixtures.js";

const _HERE = path.dirname(fileURLToPath(import.meta.url));
const TITLE_TEMPLATE = path.join(STARTER_TEMPLATES, "title-cover", "template.pptx");

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

test("roundStroke adds a round cap and a round join in schema order", () => {
  const shape = `<p:sp><p:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="000E38"/></a:solidFill><a:prstDash val="solid"/><a:tailEnd type="none"/></a:ln></p:spPr></p:sp>`;
  const out = roundStroke(shape);
  assert.match(out, /<a:ln w="12700" cap="rnd">/);
  assert.match(out, /<a:prstDash val="solid"\/><a:round\/><a:tailEnd type="none"\/>/);
});

test("roundStroke replaces an existing cap and miter join", () => {
  const shape = `<a:ln w="9525" cap="flat"><a:solidFill/><a:miter lim="800000"/></a:ln>`;
  assert.equal(roundStroke(shape), `<a:ln w="9525" cap="rnd"><a:solidFill/><a:round/></a:ln>`);
});
