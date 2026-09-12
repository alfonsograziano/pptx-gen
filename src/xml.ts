import { XMLBuilder, XMLParser } from "fast-xml-parser";

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  allowBooleanAttributes: true
});

const builder = new XMLBuilder({
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  suppressEmptyNode: true,
  format: false
});

/**
 * A node in a parsed XML tree.
 *
 * fast-xml-parser hands back plain objects whose keys are element names and
 * `@_`-prefixed attributes, and whose values are nodes, arrays of nodes, or
 * scalars — a shape that only the OOXML schema pins down. Typing the value as
 * `any` is the one deliberate escape hatch: it keeps the deep index chains in
 * `ooxml.ts` (`presentation["p:presentation"]["p:sldIdLst"]?.["p:sldId"]`)
 * readable, in one named place, instead of scattering bare `any` across every
 * call site.
 */
// biome-ignore lint/suspicious/noExplicitAny: the parsed XML tree is untyped by nature; see the note above.
export type XmlNode = Record<string, any>;

export function parseXml<T = XmlNode>(xml: string): T {
  return parser.parse(xml) as T;
}

export function buildXml(value: unknown): string {
  return builder.build(value);
}

export function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

export function escapeXml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function unescapeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
