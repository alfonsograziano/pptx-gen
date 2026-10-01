import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { cp, mkdir, mkdtemp, readFile, rm, stat, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { Presentation } from "./presentation.js";
import { CustomSlide } from "./custom-slide.js";
import { md } from "./rich-text.js";
import { PptxPackage } from "./pptx-package.js";
import { getSlideEntries } from "./ooxml.js";
import { STARTER_TEMPLATES, TINY_PNG } from "./test-fixtures.js";
import { WORKSPACE_ENV_VAR } from "./workspace.js";

const _HERE = path.dirname(fileURLToPath(import.meta.url));
const TEMPLATES = STARTER_TEMPLATES;

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

    const allText = (await Promise.all(entries.map((e) => pkg.text(`ppt/slides/slide${e.slideNumber}.xml`)))).join("");
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
    await assert.rejects(() => deck.render({ output: "deck.pptx", progress: false }), /Invalid override target/);
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
    assert.deepEqual(
      report.warnings.filter((w) => w.code === "variant-group-split"),
      []
    );
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

// A deck that names its own template library must still take logos, icons and
// design from the workspace. Passing `templateLibrary` used to switch workspace
// resolution off entirely, which pointed `assetsDir` at the pptx-gen install:
// every footer lost its logo mark, on every slide, with nothing in the report
// to say so. Seventeen real build scripts shipped that way.
test("a deck naming its own templateLibrary still takes assets from the workspace", async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptx-ws-assets-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));

  const workspace = path.join(root, "workspace");
  await mkdir(path.join(workspace, "assets"), { recursive: true });
  await writeFile(path.join(workspace, "pptx-gen.config.yml"), "version: 1\nassets: assets\n", "utf8");
  // The mark the footer asks for is present; the wordmark deliberately is not.
  // `addFooter` defaults to a light background and so asks for the dark mark,
  // while a bare `addWordmark` asks for the light one.
  await writeFile(path.join(workspace, "assets", "logo-mark-dark.png"), TINY_PNG);

  const previous = process.env[WORKSPACE_ENV_VAR];
  process.env[WORKSPACE_ENV_VAR] = workspace;
  t.after(() => {
    if (previous === undefined) delete process.env[WORKSPACE_ENV_VAR];
    else process.env[WORKSPACE_ENV_VAR] = previous;
  });

  const dir = path.join(root, "deck");
  await mkdir(dir, { recursive: true });
  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir });
  deck.addCustomSlide(
    new CustomSlide({
      name: "footer-only",
      draw: ({ slide, helpers }) => {
        helpers.addFooter(slide);
        helpers.addWordmark(slide);
      }
    })
  );

  const report = await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  // The workspace logo is in the deck, not silently dropped for the install's.
  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const media = pkg.files("ppt/media/").filter((file) => file.endsWith(".png"));
  assert.equal(media.length, 1, `expected the workspace logo in the deck, got ${JSON.stringify(media)}`);

  // And a logo that genuinely is not there says so, rather than going quiet.
  const missing = report.warnings.filter((warning) => warning.code === "logo-not-found");
  assert.deepEqual(
    missing.map((warning) => warning.target),
    ["logo-wordmark-light.png"]
  );
  assert.match(missing[0].message, /design\.yml/);
});

/** Where a relationship of `type` on `partPath` points, resolved to a package path. */
async function relatedPart(pkg: PptxPackage, partPath: string, type: string): Promise<string | undefined> {
  const relsPath = path.posix.join(path.posix.dirname(partPath), "_rels", `${path.posix.basename(partPath)}.rels`);
  if (!pkg.has(relsPath)) return undefined;
  const rel = (await pkg.text(relsPath)).match(new RegExp(`<Relationship\\b[^>]*/${type}"[^>]*>`))?.[0];
  const target = rel?.match(/Target="([^"]+)"/)?.[1];
  return target && path.posix.normalize(path.posix.join(path.posix.dirname(partPath), target));
}

/**
 * Copy a starter template into `library` as though it had been ingested from a
 * different deck, with `mutate` editing its package on the way.
 */
async function foreignTemplate(library: string, id: string, mutate: (pkg: PptxPackage) => Promise<void>) {
  const dir = path.join(library, id);
  await cp(path.join(TEMPLATES, "content-lead-bullets"), dir, { recursive: true });
  const pkg = await PptxPackage.load(path.join(dir, "template.pptx"));
  await mutate(pkg);
  await pkg.save(path.join(dir, "template.pptx"));
}

test("a template from another deck keeps its own layout, master and theme", async (t) => {
  const root = await mkdtemp(path.join(os.tmpdir(), "pptx-layout-test-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  const library = path.join(root, "templates");
  await cp(TEMPLATES, library, { recursive: true });

  // Its layout lives at a path the base deck does not have, and its master
  // hangs off a different theme: all three have to come across.
  await foreignTemplate(library, "moved-layout", async (pkg) => {
    for (const file of pkg.files()) {
      if (!file.endsWith(".rels") && file !== "[Content_Types].xml") continue;
      pkg.setText(file, (await pkg.text(file)).replaceAll("slideLayout1.xml", "slideLayout9.xml"));
    }
    await pkg.copy("ppt/slideLayouts/slideLayout1.xml", "ppt/slideLayouts/slideLayout9.xml");
    await pkg.copy("ppt/slideLayouts/_rels/slideLayout1.xml.rels", "ppt/slideLayouts/_rels/slideLayout9.xml.rels");
    pkg.remove("ppt/slideLayouts/slideLayout1.xml");
    pkg.remove("ppt/slideLayouts/_rels/slideLayout1.xml.rels");
    const theme = await pkg.text("ppt/theme/theme1.xml");
    pkg.setText("ppt/theme/theme1.xml", theme.replace('<a:clrScheme name="Office">', '<a:clrScheme name="Foreign">'));
  });
  // The dangerous case: the base has a layout at the same path, but a different
  // one. Before, the slide silently took the base's and never warned.
  await foreignTemplate(library, "same-path-layout", async (pkg) => {
    const layout = await pkg.text("ppt/slideLayouts/slideLayout1.xml");
    pkg.setText("ppt/slideLayouts/slideLayout1.xml", layout.replace('name="DEFAULT"', 'name="FOREIGN"'));
  });

  const dir = path.join(root, "deck");
  const deck = new Presentation({ templateLibrary: library, projectDir: dir });
  deck.addSlideFromTemplate({ templateName: "title-cover" });
  deck.addSlideFromTemplate({ templateName: "content-lead-bullets" });
  deck.addSlideFromTemplate({ templateName: "moved-layout" });
  deck.addSlideFromTemplate({ templateName: "same-path-layout" });
  deck.addSlideFromTemplate({ templateName: "moved-layout" });
  const report = await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  assert.deepEqual(
    report.warnings.filter((warning) => warning.code === "missing-slide-dependency"),
    []
  );

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const slides = (await getSlideEntries(pkg)).map((entry) => `ppt/slides/slide${entry.slideNumber}.xml`);
  const layouts: string[] = [];
  for (const slide of slides) layouts.push((await relatedPart(pkg, slide, "slideLayout")) ?? "none");
  const masterOf = async (layout: string) => (await relatedPart(pkg, layout, "slideMaster")) ?? "none";
  const themeOf = async (layout: string) => (await relatedPart(pkg, await masterOf(layout), "theme")) ?? "none";

  // Slices of the base deck still share its layout; nothing was copied for them.
  assert.equal(layouts[0], "ppt/slideLayouts/slideLayout1.xml");
  assert.equal(layouts[1], layouts[0]);

  // Each foreign slide sits on a layout that is byte-for-byte its own...
  const source = async (id: string, part: string) =>
    (await PptxPackage.load(path.join(library, id, "template.pptx"))).text(part);
  for (const index of [2, 3]) assert.ok(pkg.has(layouts[index]), `slide ${index + 1} has a layout`);
  assert.equal(await pkg.text(layouts[2]), await source("moved-layout", "ppt/slideLayouts/slideLayout9.xml"));
  assert.equal(await pkg.text(layouts[3]), await source("same-path-layout", "ppt/slideLayouts/slideLayout1.xml"));
  assert.notEqual(layouts[3], layouts[0]);

  // ...on its own master and theme when those differ, and on the base's when not.
  assert.notEqual(await masterOf(layouts[2]), await masterOf(layouts[0]));
  assert.equal(await pkg.text(await themeOf(layouts[2])), await source("moved-layout", "ppt/theme/theme1.xml"));
  assert.equal(await masterOf(layouts[3]), await masterOf(layouts[0]));

  // Using a foreign template twice copies its layout once.
  assert.equal(layouts[4], layouts[2]);
  const allLayouts = pkg.files("ppt/slideLayouts/").filter((file) => /slideLayout\d+\.xml$/.test(file));
  assert.equal(allLayouts.length, 3);

  // PowerPoint repairs a deck whose master and layout ids collide, or whose
  // layouts are not each listed by the master they point at.
  const presentationXml = await pkg.text("ppt/presentation.xml");
  const masters = pkg.files("ppt/slideMasters/").filter((file) => /slideMaster\d+\.xml$/.test(file));
  const ids = [...presentationXml.matchAll(/<p:sldMasterId\b[^>]*?\bid="(\d+)"/g)].map((match) => match[1]);
  const listed = new Map<string, string>();
  for (const master of masters) {
    const masterXml = await pkg.text(master);
    const rels = await pkg.text(path.posix.join("ppt/slideMasters/_rels", `${path.posix.basename(master)}.rels`));
    for (const [, id, relId] of masterXml.matchAll(/<p:sldLayoutId id="(\d+)" r:id="([^"]+)"\/>/g)) {
      ids.push(id);
      const target = rels.match(new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`))?.[1] ?? "";
      listed.set(path.posix.normalize(path.posix.join("ppt/slideMasters", target)), master);
    }
  }
  assert.equal(new Set(ids).size, ids.length, `ids are unique: ${ids.join(", ")}`);
  assert.equal(ids.filter((id) => Number(id) < 2 ** 31).length, 0);
  for (const layout of allLayouts) assert.equal(listed.get(layout), await masterOf(layout), `${layout} is listed`);
});

/** A slide's speaker notes, one line per paragraph, or undefined if it has none. */
async function notesOf(pkg: PptxPackage, slidePath: string): Promise<string | undefined> {
  const notesPath = await relatedPart(pkg, slidePath, "notesSlide");
  if (!notesPath) return undefined;
  assert.match(notesPath, /^ppt\/notesSlides\/notesSlide\d+\.xml$/);
  // Wired the way PowerPoint expects: back to its slide, and to a notes master.
  assert.equal(await relatedPart(pkg, notesPath, "slide"), slidePath);
  const master = await relatedPart(pkg, notesPath, "notesMaster");
  assert.ok(master && pkg.has(master), "the notes master exists");
  assert.ok((await pkg.text("[Content_Types].xml")).includes(`PartName="/${notesPath}"`));

  const body = ((await pkg.text(notesPath)).match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []).find((shape) =>
    /<p:ph\b[^>]*type="body"/.test(shape)
  );
  return [...(body ?? "").matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
    .map((paragraph) => [...paragraph[0].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((run) => run[1]).join(""))
    .join("\n");
}

test("speaker notes reach the deck from template and custom slides", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-notes-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir });
  deck.addSlideFromTemplate({ templateName: "title-cover", notes: "Welcome everyone.\n\nThen the agenda." });
  // pptxgenjs's own API, which the merge used to throw away.
  deck.addCustomSlide(
    new CustomSlide({ name: "draw-notes", draw: ({ slide }) => void slide.addNotes("Said in draw.") })
  );
  deck.addCustomSlide(new CustomSlide({ name: "option-notes", notes: "Set as an option.", draw: () => {} }));
  deck.addCustomSlide(new CustomSlide({ name: "no-notes", draw: () => {} }));
  await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const notes: (string | undefined)[] = [];
  for (const entry of await getSlideEntries(pkg))
    notes.push(await notesOf(pkg, `ppt/slides/slide${entry.slideNumber}.xml`));
  assert.deepEqual(notes, ["Welcome everyone.\n\nThen the agenda.", "Said in draw.", "Set as an option.", undefined]);
});

test("speaker notes survive an all-custom deck, one notes slide per slide", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-notes-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  const deck = new Presentation({ templateLibrary: TEMPLATES, projectDir: dir });
  deck.addCustomSlide(
    new CustomSlide({ name: "first", notes: "Open with the problem.\nKeep it short.", draw: () => {} })
  );
  deck.addCustomSlide(new CustomSlide({ name: "second", draw: ({ slide }) => void slide.addNotes("From draw.") }));
  await deck.render({ output: "deck.pptx", progress: false, screenshots: path.join(dir, "shots") });

  const pkg = await PptxPackage.load(path.join(dir, "deck.pptx"));
  const entries = await getSlideEntries(pkg);
  assert.equal(
    await notesOf(pkg, `ppt/slides/slide${entries[0].slideNumber}.xml`),
    "Open with the problem.\nKeep it short."
  );
  assert.equal(await notesOf(pkg, `ppt/slides/slide${entries[1].slideNumber}.xml`), "From draw.");
  // pptxgenjs had already written an empty notes slide for the first slide;
  // it is replaced, not left behind unlinked.
  const notesSlides = pkg.files("ppt/notesSlides/").filter((file) => /notesSlide\d+\.xml$/.test(file));
  assert.equal(notesSlides.length, entries.length);
});
