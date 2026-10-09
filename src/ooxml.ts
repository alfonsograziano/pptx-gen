import { createHash, randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { PptxPackage } from "./pptx-package.js";
import { fitBox, type Box, type Figure, type FigureBox, type FigureFit, type FigureRenderer } from "./figure.js";
import type { BuildWarning, SlideOverride, TemplateField, TextStyle } from "./types.js";
import { richTextToPlain } from "./rich-text.js";
import { asArray, buildXml, escapeXml, parseXml, unescapeXml, type XmlNode } from "./xml.js";
import { C, FONTS, LAYOUT } from "./design.js";
import type { AssetResolver } from "./assets.js";

const EMU_PER_IN = 914400;
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";
const SLIDE_LAYOUT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout";
const SLIDE_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster";
const NOTES_SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide";
const NOTES_MASTER_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

/** A master's list of its layouts, in any of the forms it can take. */
const SLD_LAYOUT_ID_LST = /<p:sldLayoutIdLst\/>|<p:sldLayoutIdLst>[\s\S]*?<\/p:sldLayoutIdLst>/;

/**
 * Shape name the custom-slide footer helper stamps on its page-number text box,
 * so the build can find it again and swap the literal for a live field.
 */
export const PAGE_NUMBER_SHAPE_NAME = "pptx-gen-page-number";
/** Name `addVectorIcon` gives its shapes, so the save pass can round their strokes. */
export const VECTOR_ICON_SHAPE_NAME = "pptx-gen-vector-icon";

const SLIDE_NUM_FIELD = /<a:fld\b[^>]*\btype="slidenum"/;

export type SlideEntry = {
  slideNumber: number;
  relId: string;
  target: string;
};

export async function getSlideEntries(pkg: PptxPackage): Promise<SlideEntry[]> {
  const presentation = parseXml<XmlNode>(await pkg.text("ppt/presentation.xml"));
  const rels = parseXml<XmlNode>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  const byId = new Map(relationships.map((rel: XmlNode) => [rel["@_Id"], rel]));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);

  return slideIds.map((slideId: XmlNode) => {
    const relId = slideId["@_r:id"];
    const rel = byId.get(relId);
    if (!rel) throw new Error(`Missing presentation relationship ${relId}`);
    const target = String(rel["@_Target"]);
    const match = target.match(/slides\/slide(\d+)\.xml$/);
    if (!match) throw new Error(`Unexpected slide target: ${target}`);
    return { slideNumber: Number(match[1]), relId, target };
  });
}

export async function keepOnlySlides(pkg: PptxPackage, slideNumbers: number[]): Promise<void> {
  const keep = new Set(slideNumbers);
  const entries = await getSlideEntries(pkg);
  const keepRelIds = new Set(entries.filter((entry) => keep.has(entry.slideNumber)).map((entry) => entry.relId));

  const presentation = parseXml<XmlNode>(await pkg.text("ppt/presentation.xml"));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);
  presentation["p:presentation"]["p:sldIdLst"]["p:sldId"] = slideIds.filter((slideId: XmlNode) =>
    keepRelIds.has(slideId["@_r:id"])
  );
  pkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));

  const rels = parseXml<XmlNode>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  rels.Relationships.Relationship = relationships.filter(
    (rel: XmlNode) => rel["@_Type"] !== SLIDE_REL_TYPE || keepRelIds.has(rel["@_Id"])
  );
  pkg.setText("ppt/_rels/presentation.xml.rels", withXmlHeader(buildXml(rels)));
}

/**
 * Reduce a full-deck package down to a true single-slide slice.
 *
 * The package keeps exactly one slide plus the shared support chain that every
 * slide needs (slide layouts, masters, themes, notes master, embedded fonts).
 * Every other slide, all notes slides, and any media that is no longer
 * referenced are removed, and stale content-type overrides are pruned. Because
 * the support chain is identical across all slices from the same deck, slices
 * can later be merged back together cheaply by `appendSlideFromPackage`.
 */
export async function sliceToSingleSlide(pkg: PptxPackage, ordinalSlideNumber: number): Promise<number> {
  const slides = await getSlideEntries(pkg);
  if (ordinalSlideNumber < 1 || ordinalSlideNumber > slides.length) {
    throw new Error(`Slide ${ordinalSlideNumber} is out of range. Source has ${slides.length} slide(s).`);
  }
  const keptSlideNumber = slides[ordinalSlideNumber - 1].slideNumber;

  // 1. Point the presentation at just the kept slide.
  await keepOnlySlides(pkg, [keptSlideNumber]);

  // 2. Drop every other slide part and its rels.
  for (const entry of slides) {
    if (entry.slideNumber === keptSlideNumber) continue;
    pkg.remove(`ppt/slides/slide${entry.slideNumber}.xml`);
    pkg.remove(`ppt/slides/_rels/slide${entry.slideNumber}.xml.rels`);
  }

  // 3. Drop all notes slides. The notes master stays and does not reference them.
  for (const file of pkg.files("ppt/notesSlides/")) {
    if (file.endsWith(".xml") || file.endsWith(".rels")) pkg.remove(file);
  }

  // 4. Remove the kept slide's now-dangling relationship to its notes slide.
  const keptRelsPath = `ppt/slides/_rels/slide${keptSlideNumber}.xml.rels`;
  if (pkg.has(keptRelsPath)) {
    const rels = parseXml<XmlNode>(await pkg.text(keptRelsPath));
    rels.Relationships.Relationship = asArray(rels.Relationships?.Relationship).filter(
      (rel: XmlNode) => !String(rel["@_Type"]).endsWith("/notesSlide")
    );
    pkg.setText(keptRelsPath, withXmlHeader(buildXml(rels)));
  }

  // 5. Prune media no surviving part references (the bulk of the file size).
  const referencedMedia = await collectReferencedTargets(pkg, "ppt/media/");
  for (const file of pkg.files("ppt/media/")) {
    if (!file.startsWith("ppt/media/")) continue;
    if (!referencedMedia.has(file)) pkg.remove(file);
  }

  // 6. Drop content-type overrides that point at parts we removed.
  await pruneContentTypeOverrides(pkg);

  return keptSlideNumber;
}

/**
 * Copy one slide (and any media it needs that the target lacks) from `srcPkg`
 * into `targetPkg` as a brand new slide, wiring it into the presentation and
 * content types. Returns the new slide number in the target package.
 *
 * The slide's layout is matched by content, and copied in with its master and
 * theme when the target has no equal (see `importSlideLayout`). `importLayout:
 * false` skips that and keeps whatever the target has at the same path. Fonts
 * are merged separately, by `mergeEmbeddedFonts`. Speaker notes come across as
 * a notes slide of their own when the source slide has any.
 */
export async function appendSlideFromPackage(
  targetPkg: PptxPackage,
  srcPkg: PptxPackage,
  srcSlideNumber: number,
  warnings: BuildWarning[] = [],
  options: { importLayout?: boolean } = {}
): Promise<number> {
  const importLayout = options.importLayout ?? true;
  const newSlideNumber = nextNumber(targetPkg.files("ppt/slides/"), /slide(\d+)\.xml$/);
  targetPkg.setBytes(
    `ppt/slides/slide${newSlideNumber}.xml`,
    await srcPkg.bytes(`ppt/slides/slide${srcSlideNumber}.xml`)
  );

  let srcNotesPath: string | undefined;
  const srcRelsPath = `ppt/slides/_rels/slide${srcSlideNumber}.xml.rels`;
  if (srcPkg.has(srcRelsPath)) {
    const rels = parseXml<XmlNode>(await srcPkg.text(srcRelsPath));
    // The notes slide is a part of its own, copied after the slide is wired in.
    const notesRel = asArray(rels.Relationships?.Relationship).find((rel: XmlNode) => relKind(rel) === "notesSlide");
    if (notesRel) srcNotesPath = resolveTarget(`ppt/slides/slide${srcSlideNumber}.xml`, String(notesRel["@_Target"]));
    rels.Relationships.Relationship = asArray(rels.Relationships?.Relationship).filter(
      (rel: XmlNode) => relKind(rel) !== "notesSlide"
    );
    for (const rel of asArray(rels.Relationships?.Relationship)) {
      if (String(rel["@_TargetMode"]) === "External") continue;
      const target = String(rel["@_Target"] ?? "");
      const resolved = path.posix.normalize(path.posix.join("ppt/slides", target));
      if (importLayout && relKind(rel) === "slideLayout" && srcPkg !== targetPkg && srcPkg.has(resolved)) {
        const layoutPath = await importSlideLayout(targetPkg, srcPkg, resolved);
        if (layoutPath !== resolved) rel["@_Target"] = path.posix.relative("ppt/slides", layoutPath);
      } else if (resolved.includes("/media/") && srcPkg.has(resolved)) {
        const mediaPath = await copyMediaPart(targetPkg, srcPkg, resolved);
        if (mediaPath !== resolved) rel["@_Target"] = path.posix.relative("ppt/slides", mediaPath);
      } else if (targetPkg.has(resolved)) {
      } else {
        warnings.push({
          code: "missing-slide-dependency",
          message: `Slide depends on '${resolved}' which is not present in the deck. The slide may not render correctly.`,
          slide: newSlideNumber,
          target: resolved
        });
      }
    }
    targetPkg.setText(`ppt/slides/_rels/slide${newSlideNumber}.xml.rels`, withXmlHeader(buildXml(rels)));
  }

  await addPresentationSlide(targetPkg, newSlideNumber);
  await addSlideContentType(targetPkg, newSlideNumber);

  // pptxgenjs writes a notes slide for every slide, empty unless the slide set
  // notes, so only one with something in its body is worth carrying across.
  if (srcNotesPath && srcPkg.has(srcNotesPath)) {
    const notesXml = await srcPkg.text(srcNotesPath);
    if (notesBodyText(notesXml).trim()) await attachNotesSlide(targetPkg, newSlideNumber, notesXml, warnings);
  }
  return newSlideNumber;
}

/**
 * Give a slide speaker notes, one paragraph per line of `text`, replacing any
 * it already has.
 */
export async function setSlideNotes(
  pkg: PptxPackage,
  slideNumber: number,
  text: string,
  warnings: BuildWarning[] = []
): Promise<void> {
  const paragraphs = text
    .split(/\r?\n/)
    .map((line) =>
      line
        ? `<a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
        : `<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>`
    )
    .join("");
  // The notes master supplies the geometry, so the placeholders need none.
  const notesXml = withXmlHeader(
    `<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
  );
  await attachNotesSlide(pkg, slideNumber, notesXml, warnings);
}

/** The text in a notes slide's body placeholder, which is where the notes are. */
function notesBodyText(notesXml: string): string {
  return extractShapeBlocks(notesXml)
    .filter((shapeXml) => /<p:ph\b[^>]*\btype="body"/.test(shapeXml))
    .map(getShapeText)
    .join("\n");
}

/**
 * Add `notesXml` to the package as the notes slide of `slideNumber`, tied to the
 * package's own notes master. A notes slide carries no pictures of its own in
 * anything this engine writes or copies, so its relationships are rebuilt
 * rather than carried over.
 */
async function attachNotesSlide(
  pkg: PptxPackage,
  slideNumber: number,
  notesXml: string,
  warnings: BuildWarning[]
): Promise<void> {
  const presentationRels = parseXml<XmlNode>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const masterRel = asArray(presentationRels.Relationships?.Relationship).find(
    (rel: XmlNode) => relKind(rel) === "notesMaster"
  );
  if (!masterRel) {
    warnings.push({
      code: "notes-dropped",
      message: `Slide ${slideNumber} has speaker notes, but the deck it is built on has no notes master, so they were left out. Start the deck with a template whose source deck had notes.`,
      slide: slideNumber
    });
    return;
  }
  const notesMasterPath = resolveTarget("ppt/presentation.xml", String(masterRel["@_Target"]));
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;

  // One notes slide per slide. pptxgenjs writes an empty one for every slide,
  // so an earlier one is removed outright rather than left behind unlinked.
  const slideRelsPath = relsPathOf(slidePath);
  if (pkg.has(slideRelsPath)) {
    const rels = parseXml<XmlNode>(await pkg.text(slideRelsPath));
    const relationships = asArray(rels.Relationships?.Relationship);
    const earlier = relationships.filter((rel: XmlNode) => relKind(rel) === "notesSlide");
    for (const rel of earlier) {
      const earlierPath = resolveTarget(slidePath, String(rel["@_Target"]));
      pkg.remove(earlierPath);
      pkg.remove(relsPathOf(earlierPath));
    }
    if (earlier.length > 0) {
      rels.Relationships.Relationship = relationships.filter((rel: XmlNode) => relKind(rel) !== "notesSlide");
      pkg.setText(slideRelsPath, withXmlHeader(buildXml(rels)));
      await pruneContentTypeOverrides(pkg);
    }
  }

  const notesPath = `ppt/notesSlides/notesSlide${nextNumber(pkg.files("ppt/notesSlides/"), /notesSlide(\d+)\.xml$/)}.xml`;
  pkg.setText(notesPath, notesXml);
  pkg.setText(
    relsPathOf(notesPath),
    withXmlHeader(
      buildXml({
        Relationships: {
          "@_xmlns": RELS_NS,
          Relationship: [
            {
              "@_Id": "rId1",
              "@_Type": NOTES_MASTER_REL_TYPE,
              "@_Target": path.posix.relative("ppt/notesSlides", notesMasterPath)
            },
            { "@_Id": "rId2", "@_Type": SLIDE_REL_TYPE, "@_Target": path.posix.relative("ppt/notesSlides", slidePath) }
          ]
        }
      })
    )
  );

  await addPartRelationship(pkg, slidePath, NOTES_SLIDE_REL_TYPE, path.posix.relative("ppt/slides", notesPath));
  await addOverrideContentType(
    pkg,
    notesPath,
    "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
  );
}

/**
 * Make sure `targetPkg` holds the slide layout at `srcLayoutPath` in `srcPkg`,
 * with its master and theme, and return where it lives in the target.
 *
 * A slide takes its background, its brand marks and its theme fonts from its
 * layout, master and theme, so copying the slide's XML alone is only enough
 * when the target already has the same layout. Slices of one source deck do,
 * path for path. Templates ingested from different decks do not: the same path
 * can name a different layout or none at all, and the slide then quietly
 * renders on the base deck's layout instead of its own. So layouts are matched
 * on content, and one the target has no equal of is copied in, along with its
 * master and theme when the target has no equal of those either.
 */
async function importSlideLayout(targetPkg: PptxPackage, srcPkg: PptxPackage, srcLayoutPath: string): Promise<string> {
  const srcPrints = new Map<string, string>();
  const targetPrints = new Map<string, string>();
  const existing = await findEquivalentPart(targetPkg, srcPkg, srcLayoutPath, "slideLayout", srcPrints, targetPrints);
  if (existing) return existing;

  let masterPath: string | undefined;
  for (const rel of await readRelationships(srcPkg, srcLayoutPath)) {
    if (relKind(rel) !== "slideMaster") continue;
    const srcMasterPath = resolveTarget(srcLayoutPath, String(rel["@_Target"]));
    masterPath =
      (await findEquivalentPart(targetPkg, srcPkg, srcMasterPath, "slideMaster", srcPrints, targetPrints)) ??
      (await importSlideMaster(targetPkg, srcPkg, srcMasterPath));
  }

  const layoutPath = nextPartPath(targetPkg, "slideLayout");
  await copySupportPart(targetPkg, srcPkg, srcLayoutPath, layoutPath, await srcPkg.text(srcLayoutPath), (kind) =>
    kind === "slideMaster" ? masterPath : undefined
  );
  if (masterPath) await addLayoutToMaster(targetPkg, masterPath, layoutPath);
  return layoutPath;
}

/**
 * Copy a slide master and its theme into the target as a new master with no
 * layouts yet. Layouts join it one at a time as slides need them, which keeps a
 * 200-layout source deck from dragging every one of them into the output.
 */
async function importSlideMaster(targetPkg: PptxPackage, srcPkg: PptxPackage, srcMasterPath: string): Promise<string> {
  const masterPath = nextPartPath(targetPkg, "slideMaster");
  const masterXml = (await srcPkg.text(srcMasterPath)).replace(SLD_LAYOUT_ID_LST, "<p:sldLayoutIdLst/>");

  let themePath: string | undefined;
  for (const rel of await readRelationships(srcPkg, srcMasterPath)) {
    if (relKind(rel) !== "theme") continue;
    // A theme belongs to exactly one master, so it is always copied, never shared.
    const srcThemePath = resolveTarget(srcMasterPath, String(rel["@_Target"]));
    themePath = nextPartPath(targetPkg, "theme");
    await copySupportPart(targetPkg, srcPkg, srcThemePath, themePath, await srcPkg.text(srcThemePath), () => undefined);
  }

  await copySupportPart(targetPkg, srcPkg, srcMasterPath, masterPath, masterXml, (kind) =>
    kind === "theme" ? themePath : kind === "slideLayout" ? null : undefined
  );

  const id = await nextMasterOrLayoutId(targetPkg);
  const relId = await addPresentationRelationship(
    targetPkg,
    SLIDE_MASTER_REL_TYPE,
    path.posix.relative("ppt", masterPath)
  );
  const presentation = parseXml<XmlNode>(await targetPkg.text("ppt/presentation.xml"));
  const root = presentation["p:presentation"];
  root["p:sldMasterIdLst"] ??= {};
  const list = root["p:sldMasterIdLst"];
  list["p:sldMasterId"] = [...asArray(list["p:sldMasterId"]), { "@_id": id, "@_r:id": relId }];
  targetPkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
  return masterPath;
}

/** List `layoutPath` among its master's layouts, as PowerPoint requires. */
async function addLayoutToMaster(pkg: PptxPackage, masterPath: string, layoutPath: string): Promise<void> {
  const relId = await addPartRelationship(
    pkg,
    masterPath,
    SLIDE_LAYOUT_REL_TYPE,
    path.posix.relative(path.posix.dirname(masterPath), layoutPath)
  );
  const entry = `<p:sldLayoutId id="${await nextMasterOrLayoutId(pkg)}" r:id="${relId}"/>`;
  const masterXml = await pkg.text(masterPath);
  let nextXml: string;
  if (masterXml.includes("<p:sldLayoutIdLst/>")) {
    nextXml = masterXml.replace("<p:sldLayoutIdLst/>", `<p:sldLayoutIdLst>${entry}</p:sldLayoutIdLst>`);
  } else if (masterXml.includes("</p:sldLayoutIdLst>")) {
    nextXml = masterXml.replace("</p:sldLayoutIdLst>", `${entry}</p:sldLayoutIdLst>`);
  } else {
    // The list is optional, and when present it comes straight after clrMap.
    nextXml = masterXml.replace(
      /<p:clrMap\b[^>]*\/>|<p:clrMap\b[\s\S]*?<\/p:clrMap>/,
      (clrMap) => `${clrMap}<p:sldLayoutIdLst>${entry}</p:sldLayoutIdLst>`
    );
  }
  pkg.setText(masterPath, nextXml);
}

/**
 * Master and layout ids share one space: unique across the presentation and
 * no lower than 2^31.
 */
async function nextMasterOrLayoutId(pkg: PptxPackage): Promise<number> {
  let max = 2147483647;
  const parts = ["ppt/presentation.xml", ...pkg.files("ppt/slideMasters/").filter((file) => file.endsWith(".xml"))];
  for (const part of parts) {
    for (const match of (await pkg.text(part)).matchAll(/<p:sld(?:Master|Layout)Id\b[^>]*?\bid="(\d+)"/g)) {
      max = Math.max(max, Number(match[1]));
    }
  }
  return max + 1;
}

/**
 * Find a part in the target that renders the same as `srcPath` does in its own
 * package: a layout or master whose XML, pictures, master and theme all match.
 * The same path is tried first, because slices of one deck share their parts.
 */
async function findEquivalentPart(
  targetPkg: PptxPackage,
  srcPkg: PptxPackage,
  srcPath: string,
  kind: "slideLayout" | "slideMaster",
  srcPrints: Map<string, string>,
  targetPrints: Map<string, string>
): Promise<string | undefined> {
  const wanted = await partFingerprint(srcPkg, srcPath, srcPrints);
  const pattern = new RegExp(`^ppt/${kind}s/${kind}\\d+\\.xml$`);
  const candidates = [srcPath, ...targetPkg.files(`ppt/${kind}s/`).filter((file) => pattern.test(file))];
  for (const candidate of candidates) {
    if (targetPkg.has(candidate) && (await partFingerprint(targetPkg, candidate, targetPrints)) === wanted) {
      return candidate;
    }
  }
  return undefined;
}

/**
 * Hash what a layout, master or theme looks like: its XML plus, through its
 * relationships, the bytes of its pictures and the fingerprints of its master
 * and theme. A master's own list of layouts is left out, since which layouts a
 * master offers does not change how any one of them renders — and an imported
 * master carries only the layouts the deck uses.
 */
async function partFingerprint(pkg: PptxPackage, partPath: string, memo: Map<string, string>): Promise<string> {
  const cached = memo.get(partPath);
  if (cached) return cached;

  const hash = createHash("sha1").update((await pkg.text(partPath)).replace(SLD_LAYOUT_ID_LST, ""));
  const relationships = (await readRelationships(pkg, partPath)).sort((a: XmlNode, b: XmlNode) =>
    String(a["@_Id"]).localeCompare(String(b["@_Id"]))
  );
  for (const rel of relationships) {
    const kind = relKind(rel);
    if (kind === "slideLayout") continue;
    hash.update(`\n${rel["@_Id"]} ${kind} `);
    const target = String(rel["@_Target"] ?? "");
    if (String(rel["@_TargetMode"]) === "External") {
      hash.update(target);
      continue;
    }
    const resolved = resolveTarget(partPath, target);
    if (!pkg.has(resolved)) hash.update(`missing ${resolved}`);
    else if (kind === "slideMaster" || kind === "theme") hash.update(await partFingerprint(pkg, resolved, memo));
    else hash.update(await pkg.bytes(resolved));
  }

  const print = hash.digest("hex");
  memo.set(partPath, print);
  return print;
}

/**
 * Write `xml` at `targetPath` with the relationships `srcPath` has in its own
 * package, plus its content type. `repoint` decides where a relationship of a
 * given kind now points: a path, `null` to drop it, or `undefined` for the
 * default — pictures are copied across, anything else keeps its path.
 */
async function copySupportPart(
  targetPkg: PptxPackage,
  srcPkg: PptxPackage,
  srcPath: string,
  targetPath: string,
  xml: string,
  repoint: (kind: string) => string | null | undefined
): Promise<void> {
  targetPkg.setText(targetPath, xml);

  const relationships: XmlNode[] = [];
  for (const rel of await readRelationships(srcPkg, srcPath)) {
    if (String(rel["@_TargetMode"]) === "External") {
      relationships.push(rel);
      continue;
    }
    const resolved = resolveTarget(srcPath, String(rel["@_Target"] ?? ""));
    let target = repoint(relKind(rel));
    if (target === null) continue;
    if (target === undefined) {
      target =
        resolved.includes("/media/") && srcPkg.has(resolved)
          ? await copyMediaPart(targetPkg, srcPkg, resolved)
          : resolved;
    }
    relationships.push({ ...rel, "@_Target": path.posix.relative(path.posix.dirname(targetPath), target) });
  }
  if (relationships.length > 0) {
    targetPkg.setText(
      relsPathOf(targetPath),
      withXmlHeader(buildXml({ Relationships: { "@_xmlns": RELS_NS, Relationship: relationships } }))
    );
  }

  const contentTypes = parseXml<XmlNode>(await srcPkg.text("[Content_Types].xml"));
  const override = asArray(contentTypes.Types?.Override).find((item: XmlNode) => item["@_PartName"] === `/${srcPath}`);
  if (override) await addOverrideContentType(targetPkg, targetPath, String(override["@_ContentType"]));
}

/**
 * Copy a picture into the target and return where it landed.
 *
 * Every source package numbers its own media from `image1`, so two
 * independently rendered slides routinely both carry `ppt/media/image1.png`
 * with DIFFERENT bytes. Skipping the copy because the name is taken would
 * silently point the slide at the other slide's picture, so a clashing name
 * gets a fresh one and the caller repoints its relationship at it.
 */
async function copyMediaPart(targetPkg: PptxPackage, srcPkg: PptxPackage, srcPath: string): Promise<string> {
  const bytes = await srcPkg.bytes(srcPath);
  let mediaPath = srcPath;
  if (targetPkg.has(srcPath) && !bytes.equals(await targetPkg.bytes(srcPath))) {
    const next = nextNumber(targetPkg.files("ppt/media/"), /image(\d+)\./);
    mediaPath = `ppt/media/image${next}${path.posix.extname(srcPath)}`;
  }
  if (!targetPkg.has(mediaPath)) targetPkg.setBytes(mediaPath, bytes);
  const ext = path.posix.extname(mediaPath).slice(1).toLowerCase();
  if (ext) await addDefaultContentType(targetPkg, ext, mediaContentType(ext));
  return mediaPath;
}

/** The first free `ppt/<kind>s/<kind>N.xml`, e.g. `ppt/slideLayouts/slideLayout45.xml`. */
function nextPartPath(pkg: PptxPackage, kind: "slideLayout" | "slideMaster" | "theme"): string {
  const folder = kind === "theme" ? "theme" : `${kind}s`;
  return `ppt/${folder}/${kind}${nextNumber(pkg.files(`ppt/${folder}/`), new RegExp(`${kind}(\\d+)\\.xml$`))}.xml`;
}

async function readRelationships(pkg: PptxPackage, partPath: string): Promise<XmlNode[]> {
  const relsPath = relsPathOf(partPath);
  if (!pkg.has(relsPath)) return [];
  return asArray(parseXml<XmlNode>(await pkg.text(relsPath)).Relationships?.Relationship);
}

/**
 * A relationship's kind: the last segment of its type, e.g. `slideLayout`. The
 * same in transitional and strict packages, whose type URIs differ.
 */
function relKind(rel: XmlNode): string {
  return String(rel["@_Type"] ?? "")
    .split("/")
    .pop() as string;
}

function relsPathOf(partPath: string): string {
  return path.posix.join(path.posix.dirname(partPath), "_rels", `${path.posix.basename(partPath)}.rels`);
}

function resolveTarget(partPath: string, target: string): string {
  return path.posix.normalize(path.posix.join(path.posix.dirname(partPath), target));
}

/**
 * Copy embedded font binaries that `srcPkg` carries but `targetPkg` does not.
 *
 * Each template slice is sliced from a deck that embeds its own fonts, but the
 * output deck starts from the *first* template's package. When a later slide
 * uses a font the base never embedded (e.g. "Bitter Medium"), that font is
 * absent from the output unless we carry it over. This copies the font parts,
 * adds matching presentation relationships, and extends `<p:embeddedFontLst>`
 * so the rendered deck is self-contained for every typeface its slides use.
 */
export async function mergeEmbeddedFonts(
  targetPkg: PptxPackage,
  srcPkg: PptxPackage,
  warnings: BuildWarning[] = []
): Promise<void> {
  const present = await getEmbeddedFonts(targetPkg);

  const srcPresentation = parseXml<XmlNode>(await srcPkg.text("ppt/presentation.xml"));
  const srcList = srcPresentation["p:presentation"]?.["p:embeddedFontLst"];
  const srcFonts = asArray(srcList?.["p:embeddedFont"]);
  if (srcFonts.length === 0) return;

  const srcRels = parseXml<XmlNode>(await srcPkg.text("ppt/_rels/presentation.xml.rels"));
  const srcTargetByRelId = new Map<string, string>();
  for (const rel of asArray(srcRels.Relationships?.Relationship)) {
    srcTargetByRelId.set(String(rel["@_Id"]), String(rel["@_Target"] ?? ""));
  }

  const faceTags = ["p:regular", "p:bold", "p:italic", "p:boldItalic"];
  const additions: XmlNode[] = [];

  for (const embeddedFont of srcFonts) {
    const typeface = embeddedFont["p:font"]?.["@_typeface"];
    if (!typeface || present.has(typeface)) continue;

    const merged: XmlNode = { "p:font": { "@_typeface": typeface } };
    for (const tag of faceTags) {
      const face = embeddedFont[tag];
      const srcRelId = face?.["@_r:id"];
      if (!srcRelId) continue;
      const srcTarget = srcTargetByRelId.get(String(srcRelId));
      if (!srcTarget) continue;
      const partPath = path.posix.normalize(path.posix.join("ppt", srcTarget));
      if (!srcPkg.has(partPath)) {
        warnings.push({
          code: "missing-embedded-font",
          message: `Embedded font part '${partPath}' for '${typeface}' is missing in the source template.`,
          target: typeface
        });
        continue;
      }
      if (!targetPkg.has(partPath)) targetPkg.setBytes(partPath, await srcPkg.bytes(partPath));
      const relId = await addPresentationRelationship(targetPkg, FONT_REL_TYPE, srcTarget);
      merged[tag] = { "@_r:id": relId };
    }

    if (faceTags.some((tag) => merged[tag])) {
      additions.push(merged);
      present.add(typeface);
    }
  }

  if (additions.length === 0) return;

  await addDefaultContentType(targetPkg, "fntdata", "application/x-fontdata");

  const presentation = parseXml<XmlNode>(await targetPkg.text("ppt/presentation.xml"));
  const root = presentation["p:presentation"];
  root["p:embeddedFontLst"] ??= {};
  const list = root["p:embeddedFontLst"];
  list["p:embeddedFont"] = [...asArray(list["p:embeddedFont"]), ...additions];
  targetPkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
}

async function addPresentationRelationship(pkg: PptxPackage, type: string, target: string): Promise<string> {
  const relPath = "ppt/_rels/presentation.xml.rels";
  const rels = parseXml<XmlNode>(await pkg.text(relPath));
  const relationships = asArray(rels.Relationships.Relationship);
  const relId = nextRelId(relationships);
  relationships.push({ "@_Id": relId, "@_Type": type, "@_Target": target });
  rels.Relationships.Relationship = relationships;
  pkg.setText(relPath, withXmlHeader(buildXml(rels)));
  return relId;
}

/** Collect every package part under `prefix` referenced by any surviving .rels file. */
async function collectReferencedTargets(pkg: PptxPackage, prefix: string): Promise<Set<string>> {
  const referenced = new Set<string>();
  for (const relsPath of pkg.files()) {
    if (!relsPath.endsWith(".rels")) continue;
    const ownerDir = path.posix.dirname(path.posix.dirname(relsPath));
    const rels = parseXml<XmlNode>(await pkg.text(relsPath));
    for (const rel of asArray(rels.Relationships?.Relationship)) {
      if (String(rel["@_TargetMode"]) === "External") continue;
      const target = String(rel["@_Target"] ?? "");
      const resolved = path.posix.normalize(path.posix.join(ownerDir, target));
      if (resolved.startsWith(prefix)) referenced.add(resolved);
    }
  }
  return referenced;
}

/** Remove content-type overrides whose part no longer exists in the package. */
async function pruneContentTypeOverrides(pkg: PptxPackage): Promise<void> {
  const contentTypes = parseXml<XmlNode>(await pkg.text("[Content_Types].xml"));
  contentTypes.Types.Override = asArray(contentTypes.Types?.Override).filter((override: XmlNode) =>
    pkg.has(String(override["@_PartName"]).replace(/^\//, ""))
  );
  pkg.setText("[Content_Types].xml", withXmlHeader(buildXml(contentTypes)));
}

function mediaContentType(extension: string): string {
  switch (extension) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "tiff":
      return "image/tiff";
    case "svg":
      return "image/svg+xml";
    case "emf":
      return "image/x-emf";
    case "wmf":
      return "image/x-wmf";
    default:
      return "application/octet-stream";
  }
}

export async function extractFonts(pkg: PptxPackage, slideNumber?: number): Promise<string[]> {
  const fonts = new Set<string>();
  const presentationXml = await pkg.text("ppt/presentation.xml");
  for (const match of presentationXml.matchAll(/typeface="([^"]+)"/g)) fonts.add(match[1]);
  if (slideNumber) {
    const slideXml = await pkg.text(`ppt/slides/slide${slideNumber}.xml`);
    for (const match of slideXml.matchAll(/typeface="([^"]+)"/g)) fonts.add(match[1]);
  }
  // Drop empties, ubiquitous system fonts, and theme-font references like
  // "+mn-lt" / "+mj-ea" (these resolve to the theme, not a real typeface).
  return [...fonts].filter((font) => font && !font.startsWith("+") && !["Arial", "Calibri"].includes(font)).sort();
}

/**
 * Check that every font a deck uses is either embedded in the package or a
 * common system font, and warn about any that are not.
 *
 * This is a warning, not a hard failure. A font that is embedded in the template
 * travels inside the .pptx and renders everywhere; a font that is only missing
 * locally still renders for viewers who have it, and otherwise the app
 * substitutes a fallback. Failing the build here would mean a deck cannot be
 * produced just because a design font is not installed on this machine, which is
 * exactly the friction an open tool should avoid. Install missing fonts with
 * `npm run install-fonts` if you want crisp local screenshots.
 */
export async function validateFonts(
  pkg: PptxPackage,
  expectedFonts: string[],
  warnings: BuildWarning[] = []
): Promise<void> {
  const embeddedFonts = await getEmbeddedFonts(pkg);
  const embeddedFamilies = new Set([...embeddedFonts].map(baseFontFamily));
  const isAvailable = (font: string): boolean =>
    embeddedFonts.has(font) || embeddedFamilies.has(baseFontFamily(font)) || isLikelySystemFont(font);
  const missing = expectedFonts.filter((font) => !isAvailable(font));
  if (missing.length > 0) {
    warnings.push({
      code: "font-not-embedded",
      message: `Font(s) not embedded in the deck and not known system fonts: ${missing.join(", ")}. The deck was still built; viewers without these fonts see a substitute. Run \`npm run install-fonts\` or embed them in the template PPTX for exact rendering.`
    });
  }
}

/**
 * Reduce a weight-specific typeface name to its base family so that an embedded
 * "Bitter" can satisfy a slide that asks for "Bitter Medium". Only trailing
 * weight/style words are stripped; the first word is always kept.
 */
function baseFontFamily(font: string): string {
  const weights = new Set([
    "thin",
    "extralight",
    "ultralight",
    "light",
    "regular",
    "medium",
    "semibold",
    "demibold",
    "bold",
    "extrabold",
    "ultrabold",
    "black",
    "heavy",
    "italic"
  ]);
  const words = font.trim().split(/\s+/);
  while (words.length > 1 && weights.has(words[words.length - 1].toLowerCase())) words.pop();
  return words.join(" ");
}

export async function extractTextFields(pkg: PptxPackage, slideNumber: number): Promise<TemplateField[]> {
  const slideXml = await pkg.text(`ppt/slides/slide${slideNumber}.xml`);
  const seen = new Map<string, number>();
  const textFields = extractShapeBlocks(slideXml)
    .map((shapeXml) => shapeToField(shapeXml))
    .filter((field): field is TemplateField => field !== undefined);
  // Pictures (e.g. a code panel or diagram pasted as an image) are editable too:
  // they can be swapped with the `replaceImage` override. Surface them as fields
  // so each one is discoverable and addressable by a stable id.
  const imageFields = extractPictureBlocks(slideXml)
    .map((picXml) => pictureToField(picXml))
    .filter((field): field is TemplateField => field !== undefined);
  const fields = [...textFields, ...imageFields].map((field, index) => ({
    ...field,
    // Several shapes can derive the same id from similar text (e.g. repeated
    // body copy). Field ids must be unique so each shape is addressable, so
    // collisions get a numeric suffix.
    id: uniqueFieldId(makeFieldId(field, index), seen)
  }));
  return tagPageNumberField(fields, slideXml);
}

/**
 * Tag the slide's page-number shape, so the build can turn it into a live field
 * instead of shipping whatever number the source deck happened to show. Only the
 * first match is tagged, and it gets a predictable id — otherwise the id is
 * slugified from its own text, which leaves templates carrying a field called
 * "2".
 */
function tagPageNumberField(fields: TemplateField[], slideXml: string): TemplateField[] {
  const index = fields.findIndex((field) => isPageNumberField(slideXml, field));
  if (index === -1) return fields;
  return fields.map((field, position) =>
    position === index ? { ...field, id: "page-number", role: "page-number" as const } : field
  );
}

function uniqueFieldId(baseId: string, seen: Map<string, number>): string {
  const count = seen.get(baseId) ?? 0;
  seen.set(baseId, count + 1);
  return count === 0 ? baseId : `${baseId}-${count + 1}`;
}

export async function fillSlideText(
  pkg: PptxPackage,
  slideNumber: number,
  fields: TemplateField[],
  variables: Record<string, string>,
  warnings: BuildWarning[]
): Promise<void> {
  let slideXml = await pkg.text(`ppt/slides/slide${slideNumber}.xml`);
  const fieldsById = new Map(fields.map((field) => [field.id, field]));

  for (const [id, value] of Object.entries(variables)) {
    const field = fieldsById.get(id);
    if (!field) {
      warnings.push({
        code: "unused-variable",
        message: `Variable '${id}' does not match a field`,
        slide: slideNumber,
        target: id
      });
      continue;
    }
    slideXml = replaceShape(slideXml, field, (shapeXml) => replaceShapeText(shapeXml, value));
  }

  pkg.setText(`ppt/slides/slide${slideNumber}.xml`, slideXml);
}

export type OverrideContext = {
  /** Resolves icon names, logo files, and paths relative to the deck project. */
  assets: AssetResolver;
  warnings: BuildWarning[];
  /** Shared across the build. Without it, figure overrides fall back. */
  figures?: FigureRenderer;
};

export async function applyOverrides(
  pkg: PptxPackage,
  slideNumber: number,
  fields: TemplateField[],
  overrides: SlideOverride[],
  ctx: OverrideContext
): Promise<void> {
  const { assets, warnings } = ctx;
  let slideXml = await pkg.text(`ppt/slides/slide${slideNumber}.xml`);

  for (const override of overrides) {
    if (override.op === "delete") {
      slideXml = removeShape(slideXml, override.target, fields, warnings, slideNumber);
    } else if (override.op === "hide") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) =>
        shapeXml.replace(/<p:cNvPr\b/, '<p:cNvPr hidden="1"')
      );
    } else if (override.op === "move") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) =>
        shapeXml.replace(
          /<a:off x="[^"]+" y="[^"]+"\/>/,
          `<a:off x="${inToEmu(override.x)}" y="${inToEmu(override.y)}"/>`
        )
      );
    } else if (override.op === "resize") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) =>
        shapeXml.replace(
          /<a:ext cx="[^"]+" cy="[^"]+"\/>/,
          `<a:ext cx="${inToEmu(override.w)}" cy="${inToEmu(override.h)}"/>`
        )
      );
    } else if (override.op === "styleText") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) =>
        styleShapeText(shapeXml, override)
      );
    } else if (override.op === "addText") {
      slideXml = insertShape(
        slideXml,
        createTextShape(
          override.id,
          richTextToPlain(override.text),
          override.x,
          override.y,
          override.w,
          override.h,
          override.style
        )
      );
    } else if (override.op === "addSvg" || override.op === "addIcon") {
      // addSvg names a file in the deck's own project; addIcon names an icon in
      // the icon libraries. Resolving both against the project dir, as this
      // used to, meant `{ op: "addIcon", icon: "rocket" }` worked on a custom
      // slide and failed on a cloned template slide.
      const sourcePath =
        override.op === "addSvg" ? assets.resolveProjectPath(override.path) : assets.resolveIcon(override.icon);
      const relId = await embedImagePart(pkg, slideNumber, await readFile(sourcePath), extensionOf(sourcePath, "svg"));
      slideXml = insertShape(
        slideXml,
        createPictureShape(override.id, relId, override.x, override.y, override.w, override.h)
      );
    } else if (override.op === "addImage") {
      const sourcePath = assets.resolveProjectPath(override.path);
      const relId = await embedImagePart(pkg, slideNumber, await readFile(sourcePath), extensionOf(sourcePath, "png"));
      slideXml = insertShape(
        slideXml,
        createPictureShape(override.id, relId, override.x, override.y, override.w, override.h)
      );
    } else if (override.op === "addFigure") {
      const box: Box = { x: override.x, y: override.y, w: override.w, h: override.h };
      const rendered = await renderFigure(ctx, override.figure, { w: box.w, h: box.h });
      if (!rendered) {
        // Nothing else occupies this spot, so say what belongs here.
        slideXml = insertShape(slideXml, createPlaceholderShape(override.id, override.figure.caption, box));
      } else {
        const relId = await embedImagePart(pkg, slideNumber, await readFile(rendered.pngPath), "png");
        const placed = fitBox(box, rendered.pxWidth, rendered.pxHeight, override.fit ?? "contain");
        slideXml = insertShape(
          slideXml,
          createPictureShape(override.id, relId, placed.x, placed.y, placed.w, placed.h)
        );
      }
    } else if (override.op === "replaceFigure") {
      const target = findTargetShape(slideXml, override.target, fields);
      const geometry = target ? getGeometry(target) : undefined;
      // The picture box the figure is replacing: matching its shape means the
      // figure lands exactly, with no letterboxing and nothing to re-inscribe.
      const box = geometry?.w && geometry.h ? { w: geometry.w, h: geometry.h } : undefined;
      const rendered = await renderFigure(ctx, override.figure, box);
      if (!rendered) {
        // Unlike addFigure there is already a picture in this box, and the
        // template's own image is a better stand-in than a grey placeholder.
        warnings.push({
          code: "figure-placeholder-used",
          message: `Figure '${override.figure.id}' could not be rendered, so '${override.target}' keeps the template's original image.`,
          slide: slideNumber,
          target: override.target
        });
      } else {
        const relId = await embedImagePart(pkg, slideNumber, await readFile(rendered.pngPath), "png");
        const aspect = rendered.pxWidth / rendered.pxHeight;
        slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) =>
          swapPicture(shapeXml, relId, aspect, override.fit ?? "contain", override.target, slideNumber)
        );
      }
    } else if (override.op === "replaceImage") {
      const sourcePath = assets.resolveProjectPath(override.path);
      const relId = await embedImagePart(pkg, slideNumber, await readFile(sourcePath), extensionOf(sourcePath, "png"));
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => {
        if (!/<a:blip\b[^>]*\br:embed="/.test(shapeXml)) {
          throw new Error(
            `replaceImage target '${override.target}' on slide ${slideNumber} is not an image (no <a:blip>).`
          );
        }
        return shapeXml.replace(/(<a:blip\b[^>]*\br:embed=")[^"]*(")/, `$1${relId}$2`);
      });
    }
  }

  pkg.setText(`ppt/slides/slide${slideNumber}.xml`, slideXml);
}

export async function validatePackage(filePath: string): Promise<void> {
  const pkg = await PptxPackage.load(filePath);
  await pkg.text("ppt/presentation.xml");
  await pkg.text("ppt/_rels/presentation.xml.rels");
  const slides = await getSlideEntries(pkg);
  if (slides.length === 0) throw new Error("Output PPTX has no slides");
}

/**
 * Swap the literal page number a custom slide drew for a live slide-number
 * field. Targets the text box `addFooter` named PAGE_NUMBER_SHAPE_NAME.
 */
export async function convertCustomSlidePageNumber(pkg: PptxPackage, slideNumber: number): Promise<void> {
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const slideXml = await pkg.text(slidePath);
  const nextXml = slideXml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) =>
    shapeXml.includes(`name="${PAGE_NUMBER_SHAPE_NAME}"`) ? runToSlideNumField(shapeXml) : shapeXml
  );
  if (nextXml !== slideXml) pkg.setText(slidePath, nextXml);
}

/**
 * Give every vector icon on a slide round line caps and round joins, as the
 * source SVG asks for (Lucide sets `stroke-linecap="round"`). PptxGenJS cannot
 * write either, and without them a stroke ends flat, corners go sharp, and a
 * dot (Lucide draws one as a 0.01-unit stroke, as in the "!" of triangle-alert)
 * does not render at all.
 */
export async function roundVectorIconStrokes(pkg: PptxPackage, slideNumber: number): Promise<void> {
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const slideXml = await pkg.text(slidePath);
  const nextXml = slideXml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) =>
    shapeXml.includes(`name="${VECTOR_ICON_SHAPE_NAME}"`) ? roundStroke(shapeXml) : shapeXml
  );
  if (nextXml !== slideXml) pkg.setText(slidePath, nextXml);
}

/** Set cap="rnd" on the shape outline and put `<a:round/>` where the schema wants the join. */
export function roundStroke(shapeXml: string): string {
  return shapeXml.replace(/<a:ln\b([^>]*)>([\s\S]*?)<\/a:ln>/, (match, attrs: string, inner: string) => {
    if (attrs.endsWith("/")) return match;
    const capAttrs = /\bcap="/.test(attrs) ? attrs.replace(/\bcap="[^"]*"/, 'cap="rnd"') : `${attrs} cap="rnd"`;
    // CT_LineProperties order: fill, prstDash/custDash, join (round|bevel|miter), headEnd, tailEnd, extLst.
    const withoutJoin = inner.replace(/<a:(?:round|bevel)\s*\/>|<a:miter\b[^>]*\/>/g, "");
    const at = withoutJoin.search(/<a:(?:headEnd|tailEnd|extLst)\b/);
    const joined =
      at === -1 ? `${withoutJoin}<a:round/>` : `${withoutJoin.slice(0, at)}<a:round/>${withoutJoin.slice(at)}`;
    return `<a:ln${capAttrs}>${joined}</a:ln>`;
  });
}

/**
 * The same swap for a cloned template slide. An ingested deck bakes in whatever
 * number its source slide happened to carry, and the field tagged `page-number`
 * in `fields.yml` says which shape that is.
 */
export async function convertTemplatePageNumber(
  pkg: PptxPackage,
  slideNumber: number,
  fields: TemplateField[]
): Promise<void> {
  const field = fields.find((candidate) => candidate.role === "page-number");
  if (!field) return;
  const slidePath = `ppt/slides/slide${slideNumber}.xml`;
  const slideXml = await pkg.text(slidePath);
  const nextXml = replaceShape(slideXml, field, runToSlideNumField);
  if (nextXml !== slideXml) pkg.setText(slidePath, nextXml);
}

/**
 * Settle deck-wide numbering once every slide is in its final position.
 *
 * PowerPoint renders a `slidenum` field as `firstSlideNum + (position - 1)`, so
 * to make the first slide that actually shows a footer read "1" the whole deck
 * shifts back by however many unnumbered slides (a cover, usually) come before
 * it. The literal inside each field is then rewritten to match, so a renderer
 * that ignores `<a:fld>` and paints `<a:t>` still shows the right number.
 */
export async function applySlideNumbering(pkg: PptxPackage, warnings: BuildWarning[]): Promise<void> {
  const entries = await getSlideEntries(pkg);
  const slides = await Promise.all(
    entries.map(async (entry) => ({
      entry,
      xml: await pkg.text(`ppt/slides/slide${entry.slideNumber}.xml`)
    }))
  );

  const firstNumbered = slides.findIndex((slide) => hasVisibleSlideNumber(slide.xml));
  if (firstNumbered === -1) return;

  // `firstNumbered` is 0-based, so this is `2 - position`. PowerPoint's own UI
  // only accepts 0-9999 and negative offsets are not honoured, so clamp.
  const wanted = 1 - firstNumbered;
  const firstSlideNum = Math.max(0, wanted);
  if (firstSlideNum !== wanted) {
    warnings.push({
      code: "page-numbering-clamped",
      message: `${firstNumbered} slides without a page number come before the first numbered slide. PowerPoint cannot count from below zero, so that slide reads '${firstSlideNum + firstNumbered}' rather than '1'.`
    });
  }

  await setFirstSlideNum(pkg, firstSlideNum);

  slides.forEach(({ entry, xml }, index) => {
    if (!SLIDE_NUM_FIELD.test(xml)) return;
    const nextXml = setSlideNumFallback(xml, String(firstSlideNum + index));
    if (nextXml !== xml) pkg.setText(`ppt/slides/slide${entry.slideNumber}.xml`, nextXml);
  });
}

/**
 * Turn every slide-number field back into plain text holding the number it
 * resolves to.
 *
 * This is for the screenshot copy only, never the delivered deck. LibreOffice
 * honours `slidenum` fields but ignores `firstSlideNum`, so a deck that counts
 * from 0 to skip its cover would preview one number high. Flattening first makes
 * the preview show exactly what PowerPoint will.
 */
export async function flattenSlideNumberFields(pkg: PptxPackage): Promise<void> {
  for (const entry of await getSlideEntries(pkg)) {
    const slidePath = `ppt/slides/slide${entry.slideNumber}.xml`;
    const slideXml = await pkg.text(slidePath);
    const nextXml = slideXml.replace(
      /<a:fld\b[^>]*\btype="slidenum"[^>]*>((?:(?!<\/a:fld>)[\s\S])*)<\/a:fld>/g,
      // A run holds rPr and t but not pPr, which a field is allowed to carry.
      (_match, children: string) => `<a:r>${children.replace(/<a:pPr\b[^>]*\/>|<a:pPr\b[\s\S]*?<\/a:pPr>/g, "")}</a:r>`
    );
    if (nextXml !== slideXml) pkg.setText(slidePath, nextXml);
  }
}

/**
 * Decide whether a field is the slide's page number.
 *
 * A deck this tool generated is unambiguous: the shape already holds a
 * `slidenum` field. Anything else is a guess from shape and content — a small
 * box in the bottom strip of the slide whose text is nothing but digits.
 */
function isPageNumberField(slideXml: string, field: TemplateField): boolean {
  if (field.type !== "text") return false;
  const shapeXml = extractShapeBlocks(slideXml).find((candidate) => shapeMatchesField(candidate, field));
  if (shapeXml && SLIDE_NUM_FIELD.test(shapeXml)) return true;

  const digits = field.originalText.replace(/[^0-9]/g, "");
  if (!digits || digits.length > 3) return false;
  if (/[a-z]/i.test(field.originalText)) return false;
  if (field.w === undefined || field.w > 1) return false;
  if (field.y === undefined || field.h === undefined) return false;
  return field.y + field.h > LAYOUT.height * 0.85;
}

function extractShapeBlocks(slideXml: string): string[] {
  return slideXml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? [];
}

function extractPictureBlocks(slideXml: string): string[] {
  return slideXml.match(/<p:pic>[\s\S]*?<\/p:pic>/g) ?? [];
}

function pictureToField(picXml: string): TemplateField | undefined {
  const props = picXml.match(/<p:cNvPr\b([^>]*?)\/?>/)?.[1] ?? "";
  const shapeId = props.match(/\bid="([^"]+)"/)?.[1];
  if (!shapeId) return undefined;
  const name = props.match(/\bname="([^"]+)"/)?.[1] ?? "";
  return {
    id: "",
    type: "image",
    shapeId,
    name,
    originalText: "",
    preserveStyleByDefault: false,
    ...getGeometry(picXml)
  };
}

function shapeToField(shapeXml: string): TemplateField | undefined {
  if (!shapeXml.includes("<p:txBody>")) return undefined;
  // `<p:cNvPr>` can be self-closing (`<p:cNvPr .../>`, common in Google Slides
  // exports) or a paired tag with children (`<p:cNvPr ...>...</p:cNvPr>`, emitted
  // by PptxGenJS and some PowerPoint versions). Match either form.
  const props = shapeXml.match(/<p:cNvPr\b([^>]*?)\s*\/?>/)?.[1] ?? "";
  const shapeId = props.match(/\bid="([^"]+)"/)?.[1];
  const name = props.match(/\bname="([^"]+)"/)?.[1] ?? "";
  if (!shapeId) return undefined;
  const text = getShapeText(shapeXml);
  const geometry = getGeometry(shapeXml);
  return {
    id: "",
    type: "text",
    shapeId,
    name,
    originalText: text,
    preserveStyleByDefault: true,
    ...geometry
  };
}

function getShapeText(shapeXml: string): string {
  return [...shapeXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((match) => unescapeXml(match[1])).join("");
}

function getGeometry(shapeXml: string): Pick<TemplateField, "x" | "y" | "w" | "h"> {
  const off = shapeXml.match(/<a:off x="(\d+)" y="(\d+)"\/>/);
  const ext = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
  return {
    x: off ? emuToIn(Number(off[1])) : undefined,
    y: off ? emuToIn(Number(off[2])) : undefined,
    w: ext ? emuToIn(Number(ext[1])) : undefined,
    h: ext ? emuToIn(Number(ext[2])) : undefined
  };
}

function makeFieldId(field: Pick<TemplateField, "name" | "shapeId" | "originalText">, index: number): string {
  const fromText = safeId(field.originalText).slice(0, 36);
  if (fromText) return fromText;
  const fromName = safeId(field.name).slice(0, 36);
  return fromName || `field-${field.shapeId || index + 1}`;
}

function replaceShape(slideXml: string, field: TemplateField, replacer: (shapeXml: string) => string): string {
  return slideXml.replace(/<p:sp>[\s\S]*?<\/p:sp>/g, (shapeXml) => {
    if (shapeMatchesField(shapeXml, field)) return replacer(shapeXml);
    return shapeXml;
  });
}

function replaceTargetShape(
  slideXml: string,
  target: string,
  fields: TemplateField[],
  _warnings: BuildWarning[],
  slideNumber: number,
  replacer: (shapeXml: string) => string
): string {
  let matched = false;
  const nextXml = slideXml.replace(/<p:(?:sp|pic)>[\s\S]*?<\/p:(?:sp|pic)>/g, (shapeXml) => {
    if (shapeMatchesTarget(shapeXml, target, fields)) {
      matched = true;
      return replacer(shapeXml);
    }
    return shapeXml;
  });
  if (!matched) throw new Error(`Invalid override target '${target}' on slide ${slideNumber}.`);
  return nextXml;
}

function removeShape(
  slideXml: string,
  target: string,
  fields: TemplateField[],
  warnings: BuildWarning[],
  slideNumber: number
): string {
  return replaceTargetShape(slideXml, target, fields, warnings, slideNumber, () => "");
}

function shapeMatchesField(shapeXml: string, field: TemplateField): boolean {
  return shapeXml.includes(`id="${field.shapeId}"`) || (!!field.name && shapeXml.includes(`name="${field.name}"`));
}

function shapeMatchesTarget(shapeXml: string, target: string, fields: TemplateField[]): boolean {
  const field = fields.find(
    (candidate) => candidate.id === target || candidate.shapeId === target || candidate.name === target
  );
  if (field) return shapeMatchesField(shapeXml, field);
  return shapeXml.includes(`id="${target}"`) || shapeXml.includes(`name="${target}"`);
}

function replaceShapeText(shapeXml: string, value: string): string {
  const paragraphs = value.split(/\r?\n/);
  const firstParagraph = shapeXml.match(/<a:p>[\s\S]*?<\/a:p>/)?.[0];
  if (!firstParagraph) return shapeXml;
  const newParagraphs = paragraphs.map((line) => replaceParagraphText(firstParagraph, line)).join("");
  return shapeXml.replace(/<a:p>[\s\S]*?<\/a:p>(?:\s*<a:p>[\s\S]*?<\/a:p>)*/m, newParagraphs);
}

function replaceParagraphText(paragraphXml: string, value: string): string {
  const runs = paragraphXml.match(/<a:r>[\s\S]*?<\/a:r>/g) ?? [];
  const firstExistingRun = runs[0];
  const lastExistingRun = runs[runs.length - 1];
  if (
    value.endsWith("_") &&
    firstExistingRun &&
    lastExistingRun &&
    runs.length >= 2 &&
    getShapeText(lastExistingRun) === "_"
  ) {
    const firstRun = setRunText(firstExistingRun, value.slice(0, -1));
    const lastRun = setRunText(lastExistingRun, "_");
    return paragraphXml.replace(/<a:r>[\s\S]*?<\/a:r>(?:\s*<a:r>[\s\S]*?<\/a:r>)*/m, `${firstRun}${lastRun}`);
  }
  if (!firstExistingRun) return paragraphXml;
  const firstRun = setRunText(firstExistingRun, value);
  return paragraphXml.replace(/<a:r>[\s\S]*?<\/a:r>(?:\s*<a:r>[\s\S]*?<\/a:r>)*/m, firstRun);
}

function setRunText(runXml: string, value: string): string {
  const text = escapeXml(value);
  if (runXml.includes("<a:t>")) return runXml.replace(/<a:t>[\s\S]*?<\/a:t>/, `<a:t>${text}</a:t>`);
  return runXml.replace("</a:r>", `<a:t>${text}</a:t></a:r>`);
}

function styleShapeText(shapeXml: string, style: { fontSize?: number; color?: string; fontFace?: string }): string {
  let next = shapeXml;
  const { fontSize } = style;
  if (fontSize)
    next = next.replace(/<a:rPr\b([^>]*)/g, (match) =>
      setOrReplaceAttr(match, "sz", String(Math.round(fontSize * 100)))
    );
  if (style.fontFace) next = next.replace(/typeface="[^"]+"/g, `typeface="${escapeXml(style.fontFace)}"`);
  if (style.color)
    next = next.replace(/<a:srgbClr val="[^"]+"\/>/g, `<a:srgbClr val="${style.color.replace(/^#/, "")}"/>`);
  return next;
}

function setOrReplaceAttr(tagStart: string, attr: string, value: string): string {
  if (tagStart.includes(`${attr}="`)) return tagStart.replace(new RegExp(`${attr}="[^"]*"`), `${attr}="${value}"`);
  return `${tagStart} ${attr}="${value}"`;
}

function insertShape(slideXml: string, shapeXml: string): string {
  return slideXml.replace("</p:spTree>", `${shapeXml}</p:spTree>`);
}

/** The first shape matching `target`, or undefined when nothing matches. */
function findTargetShape(slideXml: string, target: string, fields: TemplateField[]): string | undefined {
  return (slideXml.match(/<p:(?:sp|pic)>[\s\S]*?<\/p:(?:sp|pic)>/g) ?? []).find((shapeXml) =>
    shapeMatchesTarget(shapeXml, target, fields)
  );
}

/**
 * Render a figure for an override, or undefined if it could not be rendered.
 *
 * A missing browser is an environment problem and must not fail a build, so the
 * caller falls back instead. `FigureRenderer` has already recorded the warning.
 */
async function renderFigure(
  ctx: OverrideContext,
  figure: Figure,
  box?: FigureBox
): Promise<{ pngPath: string; pxWidth: number; pxHeight: number } | undefined> {
  if (!ctx.figures) return undefined;
  const result = await ctx.figures.render(figure, { box });
  return result.status === "failed" ? undefined : result;
}

/**
 * Write image bytes into the package and return a relationship id for them.
 *
 * The part is named after a hash of its content, which makes the output
 * byte-stable across rebuilds and means the same picture used on several slides
 * is stored once.
 */
async function embedImagePart(
  pkg: PptxPackage,
  slideNumber: number,
  bytes: Buffer,
  extension: string
): Promise<string> {
  const ext = extension.toLowerCase();
  const mediaName = `img-${createHash("sha1").update(bytes).digest("hex").slice(0, 10)}.${ext}`;
  const partPath = `ppt/media/${mediaName}`;
  if (!pkg.has(partPath)) pkg.setBytes(partPath, bytes);
  await addDefaultContentType(pkg, ext, mediaContentType(ext));
  return addSlideRelationship(pkg, slideNumber, IMAGE_REL_TYPE, `../media/${mediaName}`);
}

function extensionOf(filePath: string, fallback: string): string {
  return (path.extname(filePath).slice(1) || fallback).toLowerCase();
}

/**
 * Point an existing picture at new image bytes.
 *
 * Two things beyond the blip swap matter for correctness. A crop the template
 * author applied in PowerPoint lives in `<a:srcRect>` and would otherwise be
 * applied to the new image, cropping something it was never measured for. And
 * because the fill stretches, a picture whose proportions differ from the box
 * would be squashed, so the box is re-inscribed around the new aspect ratio.
 */
function swapPicture(
  shapeXml: string,
  relId: string,
  aspect: number,
  fit: FigureFit,
  target: string,
  slideNumber: number
): string {
  if (!/<a:blip\b[^>]*\br:embed="/.test(shapeXml)) {
    throw new Error(`replaceFigure target '${target}' on slide ${slideNumber} is not an image (no <a:blip>).`);
  }
  const next = shapeXml
    .replace(/(<a:blip\b[^>]*\br:embed=")[^"]*(")/, `$1${relId}$2`)
    .replace(/<a:srcRect\b[^>]*\/>/g, "");

  if (fit !== "contain") return next;
  const geometry = getGeometry(next);
  if (!geometry.x || !geometry.y || !geometry.w || !geometry.h) return next;
  const placed = fitBox({ x: geometry.x, y: geometry.y, w: geometry.w, h: geometry.h }, aspect, 1, "contain");

  // Confine the rewrite to <p:spPr>: in a <p:pic> the <p:blipFill> comes first
  // and can carry offsets of its own that must not be touched.
  return next.replace(/<p:spPr>[\s\S]*?<\/p:spPr>/, (spPr) =>
    spPr
      .replace(/<a:off x="[^"]+" y="[^"]+"\/>/, `<a:off x="${inToEmu(placed.x)}" y="${inToEmu(placed.y)}"/>`)
      .replace(/<a:ext cx="[^"]+" cy="[^"]+"\/>/, `<a:ext cx="${inToEmu(placed.w)}" cy="${inToEmu(placed.h)}"/>`)
  );
}

/**
 * A captioned stand-in for a figure that could not be rendered.
 *
 * Deliberately one `<p:sp>` so a reader can select and delete it in one click,
 * and styled to match `addImagePlaceholder` so a deck that mixes cloned and
 * custom slides does not show two different grey boxes.
 */
function createPlaceholderShape(id: string, caption: string, box: Box): string {
  const shapeId = nextRuntimeShapeId();
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="${escapeXml(id)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${inToEmu(box.x)}" y="${inToEmu(box.y)}"/><a:ext cx="${inToEmu(box.w)}" cy="${inToEmu(box.h)}"/></a:xfrm>` +
    `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 4000"/></a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${C.grey10}"/></a:solidFill>` +
    `<a:ln w="12700"><a:solidFill><a:srgbClr val="${C.grey30}"/></a:solidFill><a:prstDash val="dash"/></a:ln></p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="182880" rIns="182880"><a:noAutofit/></a:bodyPr><a:lstStyle/>` +
    `<a:p><a:pPr algn="ctr"><a:buNone/></a:pPr><a:r><a:rPr lang="en" sz="1000" i="1">` +
    `<a:solidFill><a:srgbClr val="${C.muted}"/></a:solidFill><a:latin typeface="${escapeXml(FONTS.sans)}"/></a:rPr>` +
    `<a:t>${escapeXml(caption)}</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>`
  );
}

function createTextShape(
  id: string,
  text: string,
  x: number,
  y: number,
  w: number,
  h: number,
  style: TextStyle = {}
): string {
  const shapeId = nextRuntimeShapeId();
  const color = (style.color ?? C.ink).replace(/^#/, "");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="${escapeXml(id)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${inToEmu(x)}" y="${inToEmu(y)}"/><a:ext cx="${inToEmu(w)}" cy="${inToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr lang="en" sz="${Math.round((style.fontSize ?? 10) * 100)}"${style.bold ? ` b="1"` : ""}${style.italic ? ` i="1"` : ""}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${escapeXml(style.fontFace ?? FONTS.sans)}"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>`;
}

function createPictureShape(id: string, relId: string, x: number, y: number, w: number, h: number): string {
  const shapeId = nextRuntimeShapeId();
  return `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="${escapeXml(id)}"/><p:cNvPicPr preferRelativeResize="0"/><p:nvPr/></p:nvPicPr><p:blipFill rotWithShape="1"><a:blip r:embed="${relId}"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="${inToEmu(x)}" y="${inToEmu(y)}"/><a:ext cx="${inToEmu(w)}" cy="${inToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:pic>`;
}

async function addPresentationSlide(pkg: PptxPackage, slideNumber: number): Promise<void> {
  const rels = parseXml<XmlNode>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  const relId = nextRelId(relationships);
  relationships.push({ "@_Id": relId, "@_Type": SLIDE_REL_TYPE, "@_Target": `slides/slide${slideNumber}.xml` });
  rels.Relationships.Relationship = relationships;
  pkg.setText("ppt/_rels/presentation.xml.rels", withXmlHeader(buildXml(rels)));

  const presentation = parseXml<XmlNode>(await pkg.text("ppt/presentation.xml"));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);
  const maxId = Math.max(255, ...slideIds.map((slideId: XmlNode) => Number(slideId["@_id"] ?? 255)));
  slideIds.push({ "@_id": maxId + 1, "@_r:id": relId });
  presentation["p:presentation"]["p:sldIdLst"]["p:sldId"] = slideIds;
  pkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
}

async function addSlideRelationship(
  pkg: PptxPackage,
  slideNumber: number,
  type: string,
  target: string
): Promise<string> {
  return addPartRelationship(pkg, `ppt/slides/slide${slideNumber}.xml`, type, target);
}

async function addPartRelationship(pkg: PptxPackage, partPath: string, type: string, target: string): Promise<string> {
  const relPath = relsPathOf(partPath);
  const rels = pkg.has(relPath)
    ? parseXml<XmlNode>(await pkg.text(relPath))
    : { Relationships: { "@_xmlns": RELS_NS, Relationship: [] } };
  const relationships = asArray(rels.Relationships.Relationship);
  const relId = nextRelId(relationships);
  relationships.push({ "@_Id": relId, "@_Type": type, "@_Target": target });
  rels.Relationships.Relationship = relationships;
  pkg.setText(relPath, withXmlHeader(buildXml(rels)));
  return relId;
}

async function addSlideContentType(pkg: PptxPackage, slideNumber: number): Promise<void> {
  await addOverrideContentType(
    pkg,
    `ppt/slides/slide${slideNumber}.xml`,
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
  );
}

async function addOverrideContentType(pkg: PptxPackage, partPath: string, contentType: string): Promise<void> {
  const contentTypes = parseXml<XmlNode>(await pkg.text("[Content_Types].xml"));
  const overrides = asArray(contentTypes.Types.Override);
  const partName = `/${partPath}`;
  if (!overrides.some((override: XmlNode) => override["@_PartName"] === partName)) {
    overrides.push({ "@_PartName": partName, "@_ContentType": contentType });
  }
  contentTypes.Types.Override = overrides;
  pkg.setText("[Content_Types].xml", withXmlHeader(buildXml(contentTypes)));
}

async function addDefaultContentType(pkg: PptxPackage, extension: string, contentType: string): Promise<void> {
  const contentTypes = parseXml<XmlNode>(await pkg.text("[Content_Types].xml"));
  const defaults = asArray(contentTypes.Types.Default);
  if (!defaults.some((item: XmlNode) => item["@_Extension"] === extension)) {
    defaults.push({ "@_Extension": extension, "@_ContentType": contentType });
  }
  contentTypes.Types.Default = defaults;
  pkg.setText("[Content_Types].xml", withXmlHeader(buildXml(contentTypes)));
}

function nextNumber(files: string[], pattern: RegExp): number {
  const numbers = files
    .map((file) => file.match(pattern)?.[1])
    .filter((value): value is string => Boolean(value))
    .map(Number);
  return Math.max(0, ...numbers) + 1;
}

function nextRelId(relationships: XmlNode[]): string {
  const ids = relationships
    .map((rel) => String(rel["@_Id"] ?? "").match(/^rId(\d+)$/)?.[1])
    .filter((value): value is string => Boolean(value))
    .map(Number);
  return `rId${Math.max(0, ...ids) + 1}`;
}

async function getEmbeddedFonts(pkg: PptxPackage): Promise<Set<string>> {
  const presentationXml = await pkg.text("ppt/presentation.xml");
  const embedded = new Set<string>();
  for (const match of presentationXml.matchAll(
    /<p:embeddedFont>[\s\S]*?<p:font typeface="([^"]+)"\/>[\s\S]*?<\/p:embeddedFont>/g
  )) {
    embedded.add(match[1]);
  }
  return embedded;
}

function isLikelySystemFont(font: string): boolean {
  return ["Arial", "Aptos", "Calibri", "Helvetica", "Times New Roman"].includes(font);
}

function safeId(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function withXmlHeader(xml: string): string {
  const body = xml.replace(/^(<\?xml[^>]*\?>)+/, "");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
}

function inToEmu(value: number): number {
  return Math.round(value * EMU_PER_IN);
}

function emuToIn(value: number): number {
  return Math.round((value / EMU_PER_IN) * 1000) / 1000;
}

let runtimeShapeId = 900000;

function nextRuntimeShapeId(): number {
  runtimeShapeId += 1;
  return runtimeShapeId;
}

let runtimeMediaId = 0;

function _nextRuntimeMediaId(): number {
  runtimeMediaId += 1;
  return runtimeMediaId;
}

/**
 * Rewrite a shape's last text run as a live slide-number field, keeping its
 * `<a:rPr>` verbatim so the footer holds its size, colour and typeface. The
 * stock footer is two paragraphs — a dash, then the number — and only the
 * number becomes a field.
 */
function runToSlideNumField(shapeXml: string): string {
  if (SLIDE_NUM_FIELD.test(shapeXml)) return shapeXml;
  const runs = [...shapeXml.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)];
  const lastRun = runs[runs.length - 1];
  if (!lastRun || lastRun.index === undefined) return shapeXml;

  const runXml = lastRun[0];
  const text = runXml.match(/<a:t>([\s\S]*?)<\/a:t>/)?.[1] ?? "1";
  // `<a:fld>` takes its children in the order rPr, pPr, t.
  const field = `<a:fld id="{${randomUUID().toUpperCase()}}" type="slidenum">${getRunProperties(runXml)}<a:t>${text}</a:t></a:fld>`;
  return shapeXml.slice(0, lastRun.index) + field + shapeXml.slice(lastRun.index + runXml.length);
}

/**
 * Pull a run's `<a:rPr>` out whole. It is self-closing in some exports and a
 * paired tag wrapping fills and typefaces in others, and a lazy regex would stop
 * at the first nested `/>`, so the two forms are handled apart.
 */
function getRunProperties(runXml: string): string {
  const open = runXml.match(/<a:rPr\b[^>]*>/);
  if (!open || open.index === undefined) return "";
  if (open[0].endsWith("/>")) return open[0];
  const close = runXml.indexOf("</a:rPr>", open.index);
  if (close === -1) return open[0].replace(/>$/, "/>");
  return runXml.slice(open.index, close + "</a:rPr>".length);
}

/** A slide "shows a number" only if its field sits on a shape `hide` left alone. */
function hasVisibleSlideNumber(slideXml: string): boolean {
  return extractShapeBlocks(slideXml).some(
    (shapeXml) => SLIDE_NUM_FIELD.test(shapeXml) && !/<p:cNvPr\b[^>]*\bhidden="1"/.test(shapeXml)
  );
}

function setSlideNumFallback(slideXml: string, value: string): string {
  return slideXml.replace(
    /(<a:fld\b[^>]*\btype="slidenum"[^>]*>(?:(?!<\/a:fld>)[\s\S])*?)<a:t>[\s\S]*?<\/a:t>/g,
    (_match, head: string) => `${head}<a:t>${escapeXml(value)}</a:t>`
  );
}

async function setFirstSlideNum(pkg: PptxPackage, firstSlideNum: number): Promise<void> {
  const presentation = parseXml<XmlNode>(await pkg.text("ppt/presentation.xml"));
  const root = presentation["p:presentation"];
  // 1 is the default, so leave the attribute off rather than writing a no-op.
  if (firstSlideNum === 1) delete root["@_firstSlideNum"];
  else root["@_firstSlideNum"] = firstSlideNum;
  pkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
}
