import { readFile } from "node:fs/promises";
import path from "node:path";
import { PptxPackage } from "./pptx-package.js";
import type { BuildWarning, SlideOverride, TemplateField, TextStyle } from "./types.js";
import { richTextToPlain } from "./rich-text.js";
import { asArray, buildXml, escapeXml, parseXml, unescapeXml } from "./xml.js";

const EMU_PER_IN = 914400;
const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const FONT_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font";

export type SlideEntry = {
  slideNumber: number;
  relId: string;
  target: string;
};

export async function getSlideEntries(pkg: PptxPackage): Promise<SlideEntry[]> {
  const presentation = parseXml<any>(await pkg.text("ppt/presentation.xml"));
  const rels = parseXml<any>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  const byId = new Map(relationships.map((rel: any) => [rel["@_Id"], rel]));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);

  return slideIds.map((slideId: any) => {
    const relId = slideId["@_r:id"];
    const rel = byId.get(relId);
    if (!rel) throw new Error(`Missing presentation relationship ${relId}`);
    const target = String(rel["@_Target"]);
    const match = target.match(/slides\/slide(\d+)\.xml$/);
    if (!match) throw new Error(`Unexpected slide target: ${target}`);
    return { slideNumber: Number(match[1]), relId, target };
  });
}

export async function duplicateSlide(pkg: PptxPackage, sourceSlideNumber: number): Promise<number> {
  const nextSlideNumber = nextNumber(pkg.files("ppt/slides/"), /slide(\d+)\.xml$/);
  await pkg.copy(`ppt/slides/slide${sourceSlideNumber}.xml`, `ppt/slides/slide${nextSlideNumber}.xml`);
  await pkg.copy(`ppt/slides/_rels/slide${sourceSlideNumber}.xml.rels`, `ppt/slides/_rels/slide${nextSlideNumber}.xml.rels`);
  await addPresentationSlide(pkg, nextSlideNumber);
  await addSlideContentType(pkg, nextSlideNumber);
  return nextSlideNumber;
}

export async function keepOnlySlides(pkg: PptxPackage, slideNumbers: number[]): Promise<void> {
  const keep = new Set(slideNumbers);
  const entries = await getSlideEntries(pkg);
  const keepRelIds = new Set(entries.filter((entry) => keep.has(entry.slideNumber)).map((entry) => entry.relId));

  const presentation = parseXml<any>(await pkg.text("ppt/presentation.xml"));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);
  presentation["p:presentation"]["p:sldIdLst"]["p:sldId"] = slideIds.filter((slideId: any) => keepRelIds.has(slideId["@_r:id"]));
  pkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));

  const rels = parseXml<any>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  rels.Relationships.Relationship = relationships.filter((rel: any) => (
    rel["@_Type"] !== SLIDE_REL_TYPE || keepRelIds.has(rel["@_Id"])
  ));
  pkg.setText("ppt/_rels/presentation.xml.rels", withXmlHeader(buildXml(rels)));
}

export async function keepOnlySlideByOrdinal(pkg: PptxPackage, ordinalSlideNumber: number): Promise<number> {
  const slides = await getSlideEntries(pkg);
  if (ordinalSlideNumber < 1 || ordinalSlideNumber > slides.length) {
    throw new Error(`Slide ${ordinalSlideNumber} is out of range. Source has ${slides.length} slide(s).`);
  }

  const keptSlideNumber = slides[ordinalSlideNumber - 1].slideNumber;
  await keepOnlySlides(pkg, [keptSlideNumber]);
  return keptSlideNumber;
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
    const rels = parseXml<any>(await pkg.text(keptRelsPath));
    rels.Relationships.Relationship = asArray(rels.Relationships?.Relationship)
      .filter((rel: any) => !String(rel["@_Type"]).endsWith("/notesSlide"));
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
 * Slide layouts, masters, themes, and fonts are assumed to already exist in the
 * target (true for slices from the same source deck), so only the slide part,
 * its rels, and referenced media are copied.
 */
export async function appendSlideFromPackage(
  targetPkg: PptxPackage,
  srcPkg: PptxPackage,
  srcSlideNumber: number,
  warnings: BuildWarning[] = []
): Promise<number> {
  const newSlideNumber = nextNumber(targetPkg.files("ppt/slides/"), /slide(\d+)\.xml$/);
  targetPkg.setBytes(`ppt/slides/slide${newSlideNumber}.xml`, await srcPkg.bytes(`ppt/slides/slide${srcSlideNumber}.xml`));

  const srcRelsPath = `ppt/slides/_rels/slide${srcSlideNumber}.xml.rels`;
  if (srcPkg.has(srcRelsPath)) {
    const rels = parseXml<any>(await srcPkg.text(srcRelsPath));
    rels.Relationships.Relationship = asArray(rels.Relationships?.Relationship)
      .filter((rel: any) => !String(rel["@_Type"]).endsWith("/notesSlide"));
    for (const rel of asArray(rels.Relationships?.Relationship)) {
      if (String(rel["@_TargetMode"]) === "External") continue;
      const target = String(rel["@_Target"] ?? "");
      const resolved = path.posix.normalize(path.posix.join("ppt/slides", target));
      if (targetPkg.has(resolved)) continue;
      if (resolved.includes("/media/") && srcPkg.has(resolved)) {
        targetPkg.setBytes(resolved, await srcPkg.bytes(resolved));
        const ext = path.posix.extname(resolved).slice(1).toLowerCase();
        if (ext) await addDefaultContentType(targetPkg, ext, mediaContentType(ext));
      } else if (!targetPkg.has(resolved)) {
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
  return newSlideNumber;
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

  const srcPresentation = parseXml<any>(await srcPkg.text("ppt/presentation.xml"));
  const srcList = srcPresentation["p:presentation"]?.["p:embeddedFontLst"];
  const srcFonts = asArray(srcList?.["p:embeddedFont"]);
  if (srcFonts.length === 0) return;

  const srcRels = parseXml<any>(await srcPkg.text("ppt/_rels/presentation.xml.rels"));
  const srcTargetByRelId = new Map<string, string>();
  for (const rel of asArray(srcRels.Relationships?.Relationship)) {
    srcTargetByRelId.set(String(rel["@_Id"]), String(rel["@_Target"] ?? ""));
  }

  const faceTags = ["p:regular", "p:bold", "p:italic", "p:boldItalic"];
  const additions: any[] = [];

  for (const embeddedFont of srcFonts) {
    const typeface = embeddedFont["p:font"]?.["@_typeface"];
    if (!typeface || present.has(typeface)) continue;

    const merged: any = { "p:font": { "@_typeface": typeface } };
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

  const presentation = parseXml<any>(await targetPkg.text("ppt/presentation.xml"));
  const root = presentation["p:presentation"];
  const list = root["p:embeddedFontLst"] ?? (root["p:embeddedFontLst"] = {});
  list["p:embeddedFont"] = [...asArray(list["p:embeddedFont"]), ...additions];
  targetPkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
}

async function addPresentationRelationship(pkg: PptxPackage, type: string, target: string): Promise<string> {
  const relPath = "ppt/_rels/presentation.xml.rels";
  const rels = parseXml<any>(await pkg.text(relPath));
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
    const rels = parseXml<any>(await pkg.text(relsPath));
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
  const contentTypes = parseXml<any>(await pkg.text("[Content_Types].xml"));
  contentTypes.Types.Override = asArray(contentTypes.Types?.Override)
    .filter((override: any) => pkg.has(String(override["@_PartName"]).replace(/^\//, "")));
  pkg.setText("[Content_Types].xml", withXmlHeader(buildXml(contentTypes)));
}

function mediaContentType(extension: string): string {
  switch (extension) {
    case "jpg":
    case "jpeg": return "image/jpeg";
    case "png": return "image/png";
    case "gif": return "image/gif";
    case "bmp": return "image/bmp";
    case "tiff": return "image/tiff";
    case "svg": return "image/svg+xml";
    case "emf": return "image/x-emf";
    case "wmf": return "image/x-wmf";
    default: return "application/octet-stream";
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
  return [...fonts].filter((font) => font && !["Arial", "Calibri"].includes(font)).sort();
}

export async function validateFonts(pkg: PptxPackage, expectedFonts: string[]): Promise<void> {
  const embeddedFonts = await getEmbeddedFonts(pkg);
  const embeddedFamilies = new Set([...embeddedFonts].map(baseFontFamily));
  const isAvailable = (font: string): boolean =>
    embeddedFonts.has(font) ||
    embeddedFamilies.has(baseFontFamily(font)) ||
    isLikelySystemFont(font);
  const missing = expectedFonts.filter((font) => !isAvailable(font));
  if (missing.length > 0) {
    throw new Error(`Missing required font(s): ${missing.join(", ")}. Run \`npm run ensure-fonts\` to install them, or embed them in the template PPTX.`);
  }
}

/**
 * Reduce a weight-specific typeface name to its base family so that an embedded
 * "Bitter" can satisfy a slide that asks for "Bitter Medium". Only trailing
 * weight/style words are stripped; the first word is always kept.
 */
function baseFontFamily(font: string): string {
  const weights = new Set([
    "thin", "extralight", "ultralight", "light", "regular", "medium",
    "semibold", "demibold", "bold", "extrabold", "ultrabold", "black", "heavy", "italic"
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
  return [...textFields, ...imageFields].map((field, index) => ({
    ...field,
    // Several shapes can derive the same id from similar text (e.g. repeated
    // body copy). Field ids must be unique so each shape is addressable, so
    // collisions get a numeric suffix.
    id: uniqueFieldId(makeFieldId(field, index), seen)
  }));
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
      warnings.push({ code: "unused-variable", message: `Variable '${id}' does not match a field`, slide: slideNumber, target: id });
      continue;
    }
    slideXml = replaceShape(slideXml, field, (shapeXml) => replaceShapeText(shapeXml, value));
  }

  pkg.setText(`ppt/slides/slide${slideNumber}.xml`, slideXml);
}

export async function applyOverrides(
  pkg: PptxPackage,
  slideNumber: number,
  fields: TemplateField[],
  overrides: SlideOverride[],
  rootDir: string,
  warnings: BuildWarning[]
): Promise<void> {
  let slideXml = await pkg.text(`ppt/slides/slide${slideNumber}.xml`);

  for (const override of overrides) {
    if (override.op === "delete") {
      slideXml = removeShape(slideXml, override.target, fields, warnings, slideNumber);
    } else if (override.op === "hide") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => (
        shapeXml.replace(/<p:cNvPr\b/, "<p:cNvPr hidden=\"1\"")
      ));
    } else if (override.op === "move") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => (
        shapeXml.replace(/<a:off x="[^"]+" y="[^"]+"\/>/, `<a:off x="${inToEmu(override.x)}" y="${inToEmu(override.y)}"/>`)
      ));
    } else if (override.op === "resize") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => (
        shapeXml.replace(/<a:ext cx="[^"]+" cy="[^"]+"\/>/, `<a:ext cx="${inToEmu(override.w)}" cy="${inToEmu(override.h)}"/>`)
      ));
    } else if (override.op === "styleText") {
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => styleShapeText(shapeXml, override));
    } else if (override.op === "addText") {
      slideXml = insertShape(slideXml, createTextShape(override.id, richTextToPlain(override.text), override.x, override.y, override.w, override.h, override.style));
    } else if (override.op === "addSvg" || override.op === "addIcon") {
      const sourcePath = path.resolve(rootDir, override.op === "addSvg" ? override.path : override.icon);
      const svg = await readFile(sourcePath, "utf8");
      const mediaName = `ppt/media/${safeId(override.id)}-${Date.now()}.svg`;
      pkg.setBytes(mediaName, svg);
      await addDefaultContentType(pkg, "svg", "image/svg+xml");
      const relId = await addSlideRelationship(pkg, slideNumber, IMAGE_REL_TYPE, `../media/${path.basename(mediaName)}`);
      slideXml = insertShape(slideXml, createPictureShape(override.id, relId, override.x, override.y, override.w, override.h));
    } else if (override.op === "replaceImage") {
      const sourcePath = path.resolve(rootDir, override.path);
      const data = await readFile(sourcePath);
      const ext = (path.extname(sourcePath).slice(1) || "png").toLowerCase();
      const mediaName = `${safeId(override.target)}-${slideNumber}-${nextRuntimeMediaId()}.${ext}`;
      pkg.setBytes(`ppt/media/${mediaName}`, data);
      await addDefaultContentType(pkg, ext, mediaContentType(ext));
      const relId = await addSlideRelationship(pkg, slideNumber, IMAGE_REL_TYPE, `../media/${mediaName}`);
      slideXml = replaceTargetShape(slideXml, override.target, fields, warnings, slideNumber, (shapeXml) => {
        if (!/<a:blip\b[^>]*\br:embed="/.test(shapeXml)) {
          throw new Error(`replaceImage target '${override.target}' on slide ${slideNumber} is not an image (no <a:blip>).`);
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
  const props = shapeXml.match(/<p:cNvPr\b([^>]*)\/>/)?.[1] ?? "";
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
  return [...shapeXml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
    .map((match) => unescapeXml(match[1]))
    .join("");
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
  warnings: BuildWarning[],
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

function removeShape(slideXml: string, target: string, fields: TemplateField[], warnings: BuildWarning[], slideNumber: number): string {
  return replaceTargetShape(slideXml, target, fields, warnings, slideNumber, () => "");
}

function shapeMatchesField(shapeXml: string, field: TemplateField): boolean {
  return shapeXml.includes(`id="${field.shapeId}"`) || (!!field.name && shapeXml.includes(`name="${field.name}"`));
}

function shapeMatchesTarget(shapeXml: string, target: string, fields: TemplateField[]): boolean {
  const field = fields.find((candidate) => candidate.id === target || candidate.shapeId === target || candidate.name === target);
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
  if (value.endsWith("_") && firstExistingRun && lastExistingRun && runs.length >= 2 && getShapeText(lastExistingRun) === "_") {
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
  if (style.fontSize) next = next.replace(/<a:rPr\b([^>]*)/g, (match) => setOrReplaceAttr(match, "sz", String(Math.round(style.fontSize! * 100))));
  if (style.fontFace) next = next.replace(/typeface="[^"]+"/g, `typeface="${escapeXml(style.fontFace)}"`);
  if (style.color) next = next.replace(/<a:srgbClr val="[^"]+"\/>/g, `<a:srgbClr val="${style.color.replace(/^#/, "")}"/>`);
  return next;
}

function setOrReplaceAttr(tagStart: string, attr: string, value: string): string {
  if (tagStart.includes(`${attr}="`)) return tagStart.replace(new RegExp(`${attr}="[^"]*"`), `${attr}="${value}"`);
  return `${tagStart} ${attr}="${value}"`;
}

function insertShape(slideXml: string, shapeXml: string): string {
  return slideXml.replace("</p:spTree>", `${shapeXml}</p:spTree>`);
}

function createTextShape(id: string, text: string, x: number, y: number, w: number, h: number, style: TextStyle = {}): string {
  const shapeId = nextRuntimeShapeId();
  const color = (style.color ?? "000E38").replace(/^#/, "");
  return `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="${escapeXml(id)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${inToEmu(x)}" y="${inToEmu(y)}"/><a:ext cx="${inToEmu(w)}" cy="${inToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l"><a:buNone/></a:pPr><a:r><a:rPr lang="en" sz="${Math.round((style.fontSize ?? 10) * 100)}"${style.bold ? ` b="1"` : ""}${style.italic ? ` i="1"` : ""}><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="${escapeXml(style.fontFace ?? "Inter")}"/></a:rPr><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>`;
}

function createPictureShape(id: string, relId: string, x: number, y: number, w: number, h: number): string {
  const shapeId = nextRuntimeShapeId();
  return `<p:pic><p:nvPicPr><p:cNvPr id="${shapeId}" name="${escapeXml(id)}"/><p:cNvPicPr preferRelativeResize="0"/><p:nvPr/></p:nvPicPr><p:blipFill rotWithShape="1"><a:blip r:embed="${relId}"/><a:stretch/></p:blipFill><p:spPr><a:xfrm><a:off x="${inToEmu(x)}" y="${inToEmu(y)}"/><a:ext cx="${inToEmu(w)}" cy="${inToEmu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr></p:pic>`;
}

async function addPresentationSlide(pkg: PptxPackage, slideNumber: number): Promise<void> {
  const rels = parseXml<any>(await pkg.text("ppt/_rels/presentation.xml.rels"));
  const relationships = asArray(rels.Relationships.Relationship);
  const relId = nextRelId(relationships);
  relationships.push({ "@_Id": relId, "@_Type": SLIDE_REL_TYPE, "@_Target": `slides/slide${slideNumber}.xml` });
  rels.Relationships.Relationship = relationships;
  pkg.setText("ppt/_rels/presentation.xml.rels", withXmlHeader(buildXml(rels)));

  const presentation = parseXml<any>(await pkg.text("ppt/presentation.xml"));
  const slideIds = asArray(presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]);
  const maxId = Math.max(255, ...slideIds.map((slideId: any) => Number(slideId["@_id"] ?? 255)));
  slideIds.push({ "@_id": maxId + 1, "@_r:id": relId });
  presentation["p:presentation"]["p:sldIdLst"]["p:sldId"] = slideIds;
  pkg.setText("ppt/presentation.xml", withXmlHeader(buildXml(presentation)));
}

async function addSlideRelationship(pkg: PptxPackage, slideNumber: number, type: string, target: string): Promise<string> {
  const relPath = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
  const rels = pkg.has(relPath)
    ? parseXml<any>(await pkg.text(relPath))
    : { Relationships: { "@_xmlns": "http://schemas.openxmlformats.org/package/2006/relationships", Relationship: [] } };
  const relationships = asArray(rels.Relationships.Relationship);
  const relId = nextRelId(relationships);
  relationships.push({ "@_Id": relId, "@_Type": type, "@_Target": target });
  rels.Relationships.Relationship = relationships;
  pkg.setText(relPath, withXmlHeader(buildXml(rels)));
  return relId;
}

async function addSlideContentType(pkg: PptxPackage, slideNumber: number): Promise<void> {
  const contentTypes = parseXml<any>(await pkg.text("[Content_Types].xml"));
  const overrides = asArray(contentTypes.Types.Override);
  const partName = `/ppt/slides/slide${slideNumber}.xml`;
  if (!overrides.some((override: any) => override["@_PartName"] === partName)) {
    overrides.push({
      "@_PartName": partName,
      "@_ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    });
  }
  contentTypes.Types.Override = overrides;
  pkg.setText("[Content_Types].xml", withXmlHeader(buildXml(contentTypes)));
}

async function addDefaultContentType(pkg: PptxPackage, extension: string, contentType: string): Promise<void> {
  const contentTypes = parseXml<any>(await pkg.text("[Content_Types].xml"));
  const defaults = asArray(contentTypes.Types.Default);
  if (!defaults.some((item: any) => item["@_Extension"] === extension)) {
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

function nextRelId(relationships: any[]): string {
  const ids = relationships
    .map((rel) => String(rel["@_Id"] ?? "").match(/^rId(\d+)$/)?.[1])
    .filter((value): value is string => Boolean(value))
    .map(Number);
  return `rId${Math.max(0, ...ids) + 1}`;
}

async function getEmbeddedFonts(pkg: PptxPackage): Promise<Set<string>> {
  const presentationXml = await pkg.text("ppt/presentation.xml");
  const embedded = new Set<string>();
  for (const match of presentationXml.matchAll(/<p:embeddedFont>[\s\S]*?<p:font typeface="([^"]+)"\/>[\s\S]*?<\/p:embeddedFont>/g)) {
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

function nextRuntimeMediaId(): number {
  runtimeMediaId += 1;
  return runtimeMediaId;
}
