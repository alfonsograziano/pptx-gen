import type { MarkdownText, PrimitiveRichText } from "./types.js";

export function md(value: string): MarkdownText {
  return { kind: "markdown", value };
}

export function richTextToPlain(value: PrimitiveRichText): string {
  if (typeof value === "string") return value;
  return markdownToPlain(value.value);
}

export function markdownToPlain(value: string): string {
  return value
    .trim()
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/__(.*?)__/g, "$1")
    .replace(/\*(.*?)\*/g, "$1")
    .replace(/_(.*?)_/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    // Drop the list marker itself. Each line becomes a separate paragraph and
    // inherits the template paragraph's own bullet styling, so adding a literal
    // "-" here would either show a stray dash or double up an existing bullet.
    .replace(/^\s*[-*+]\s+/gm, "")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1");
}
