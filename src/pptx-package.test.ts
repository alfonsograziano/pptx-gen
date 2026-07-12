import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, rm } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { PptxPackage } from "./pptx-package.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATE = path.resolve(HERE, "..", "templates", "title-cover", "template.pptx");

test("load exposes package parts", async () => {
  const pkg = await PptxPackage.load(TEMPLATE);
  assert.ok(pkg.has("ppt/presentation.xml"));
  assert.ok(!pkg.has("ppt/does-not-exist.xml"));
  const xml = await pkg.text("ppt/presentation.xml");
  assert.match(xml, /<p:presentation/);
  assert.ok(pkg.files("ppt/slides/").some((f) => /slide\d+\.xml$/.test(f)));
});

test("setText then save then reload round-trips an edit", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-pkg-test-"));
  try {
    const pkg = await PptxPackage.load(TEMPLATE);
    pkg.setText("docProps/custom-marker.xml", "<marker/>");
    const out = path.join(dir, "roundtrip.pptx");
    await pkg.save(out);

    const reloaded = await PptxPackage.load(out);
    assert.ok(reloaded.has("docProps/custom-marker.xml"));
    assert.equal(await reloaded.text("docProps/custom-marker.xml"), "<marker/>");
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
