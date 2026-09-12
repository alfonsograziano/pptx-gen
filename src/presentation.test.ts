import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { md } from "./rich-text.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATES = path.resolve(HERE, "..", "starter", "templates");

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

test("variant groups are numbered and grouped in the report", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });

    deck.addSlideFromTemplate({
      templateName: "title-cover",
      variables: { "your-presentation-title-goes-here": "Cover" }
    });

    // Three takes on one concept. Real variants are different layouts; here the
    // same template stands in so the test stays fast and needs no rendering.
    const leads = ["A timeline of the work", "The work, as cards", "The work, numbered"];
    for (const [index, lead] of leads.entries()) {
      deck.addSlideFromTemplate({
        templateName: "content-lead-bullets",
        group: "agenda",
        variables: {
          "section-header": `Agenda ${index + 1}_`,
          "a-lead-statement-that-frames-the-thr": lead
        }
      });
    }

    deck.addSlideFromTemplate({
      templateName: "title-cover",
      variables: { "your-presentation-title-goes-here": "Thanks" }
    });

    const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

    // Two concepts, the first with three variants, render as five slides.
    assert.equal(report.slidesBuilt, 5);
    assert.equal(report.slides.length, 5);
    const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
    assert.equal((await getSlideEntries(pkg)).length, 5);

    assert.equal(report.slides[0].group, undefined);
    assert.equal(report.slides[4].group, undefined);
    assert.deepEqual(
      report.slides.slice(1, 4).map((s) => [s.index, s.group, s.variant, s.variantCount]),
      [
        [2, "agenda", 1, 3],
        [3, "agenda", 2, 3],
        [4, "agenda", 3, 3]
      ]
    );

    const reportMd = await readFile(path.join(dir, "report.md"), "utf8");
    assert.match(reportMd, /- Variant groups: agenda \(3\)/);
    assert.match(reportMd, /- Variant group `agenda` — 3 variants, slides 2-4\. Pick one:/);
    assert.match(reportMd, /variant 2 of 3/);
    assert.match(reportMd, /- 1\. title-cover \(template\)/);

    const allowed = new Set(["font-not-embedded", "screenshots-skipped"]);
    const unexpected = report.warnings.filter((w) => !allowed.has(w.code));
    assert.deepEqual(unexpected, [], `unexpected warnings: ${JSON.stringify(unexpected)}`);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a non-consecutive variant group warns but still builds", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });

    deck.addSlideFromTemplate({ templateName: "title-cover", group: "cover", variables: { "overline-label": "One" } });
    deck.addSlideFromTemplate({ templateName: "content-lead-bullets", variables: { "section-header": "Interloper_" } });
    deck.addSlideFromTemplate({ templateName: "title-cover", group: "cover", variables: { "overline-label": "Two" } });

    const report = await deck.render({ output: "deck.pptx", progress: false });

    assert.equal(report.slidesBuilt, 3);
    const split = report.warnings.filter((w) => w.code === "variant-group-split");
    assert.equal(split.length, 1);
    assert.equal(split[0].target, "cover");
    assert.match(split[0].message, /positions 1, 3/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("a group with a single member is not a variant group", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });
    deck.addSlideFromTemplate({ templateName: "title-cover", group: "solo", variables: { "overline-label": "x" } });

    const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

    assert.equal(report.slides[0].group, undefined);
    assert.equal(report.slides[0].variant, undefined);
    assert.equal(report.slides[0].variantCount, undefined);

    const reportMd = await readFile(path.join(dir, "report.md"), "utf8");
    assert.match(reportMd, /- Variant groups: none/);
    assert.doesNotMatch(reportMd, /Variant group `/);
    assert.deepEqual(report.warnings.filter((w) => w.code === "variant-group-split"), []);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});

test("variant groups work on the all-custom render path", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-build-test-"));
  try {
    const deck = new Presentation({ title: "Test deck", templateLibrary: TEMPLATES, projectDir: dir });

    const variant = (name: string, text: string) =>
      new CustomSlide({
        name,
        group: "opening",
        draw: ({ slide }) => {
          slide.addText(text, { x: 0.8, y: 2.2, w: 8.4, h: 1.0, fontSize: 32 });
        }
      });

    deck.addCustomSlide(variant("opening-statement", "One big claim"));
    deck.addCustomSlide(variant("opening-question", "A question, then the claim"));

    const report = await deck.render({ output: "deck.pptx", report: "report.md", progress: false });

    assert.equal(report.slidesBuilt, 2);
    assert.deepEqual(
      report.slides.map((s) => [s.kind, s.name, s.group, s.variant, s.variantCount]),
      [
        ["custom", "opening-statement", "opening", 1, 2],
        ["custom", "opening-question", "opening", 2, 2]
      ]
    );

    const reportMd = await readFile(path.join(dir, "report.md"), "utf8");
    assert.match(reportMd, /- Variant group `opening` — 2 variants, slides 1-2\. Pick one:/);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
});
