import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, rm, stat } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { Presentation } from "./presentation.js";
import { md } from "./rich-text.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATES = path.resolve(HERE, "..", "templates");

test("builds a two-slide deck from templates with replaced text", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });

    deck.addSlideFromTemplate({
      templateName: "title-cover",
      variables: {
        "overline-label": "Proposal",
        "your-presentation-title-goes-here": "A practical path to production",
        "a-short-subtitle-that-sets-up-the-st": "From experiments to safe delivery."
      }
    });

    deck.addSlideFromTemplate({
      templateName: "content-lead-bullets",
      variables: {
        "section-header": "Why this works_",
        "a-lead-statement-that-frames-the-thr": md("Three moves get you there."),
        "first-supporting-point-that-backs-up": "Start with one workflow\nAdd evaluation gates\nTrack cost and quality"
      }
    });

    const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

    // The .pptx exists and is a valid package with exactly the slides we built.
    const outputPath = path.join(dir, "deck.pptx");
    assert.ok((await stat(outputPath)).size > 0);
    const pkg = await PptxPackage.load(outputPath);
    const entries = await getSlideEntries(pkg);
    assert.equal(entries.length, 2);

    const allText = (
      await Promise.all(entries.map((e) => pkg.text(`ppt/slides/slide${e.slideNumber}.xml`)))
    ).join("");
    for (const expected of [
      "A practical path to production",
      "From experiments to safe delivery.",
      "Why this works",
      "Add evaluation gates"
    ]) {
      assert.ok(allText.includes(expected), `expected replaced text: ${expected}`);
    }

    assert.equal(report.slidesBuilt, 2);
    assert.deepEqual(report.templatesUsed, ["title-cover", "content-lead-bullets"]);

    // The only warnings allowed are the environment-dependent, non-fatal ones.
    const allowed = new Set(["font-not-embedded", "screenshots-skipped"]);
    const unexpected = report.warnings.filter((w) => !allowed.has(w.code));
    assert.deepEqual(unexpected, [], `unexpected warnings: ${JSON.stringify(unexpected)}`);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("an unknown override target fails the build", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });
    deck.addSlideFromTemplate({
      templateName: "title-cover",
      variables: { "overline-label": "x" },
      overrides: [{ op: "delete", target: "no-such-shape" }]
    });
    await assert.rejects(
      () => deck.render({ output: "deck.pptx", progress: false }),
      /Invalid override target/
    );
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
