import { test } from "node:test";
import assert from "node:assert/strict";
import { md, markdownToPlain, richTextToPlain } from "./rich-text.js";

test("md wraps a string as markdown rich text", () => {
  assert.deepEqual(md("hello"), { kind: "markdown", value: "hello" });
});

test("richTextToPlain returns plain strings unchanged", () => {
  assert.equal(richTextToPlain("just text"), "just text");
});

test("richTextToPlain flattens markdown emphasis", () => {
  assert.equal(richTextToPlain(md("A **bold** and _italic_ line")), "A bold and italic line");
});

test("markdownToPlain strips list markers but keeps line breaks", () => {
  assert.equal(markdownToPlain("- one\n- two\n- three"), "one\ntwo\nthree");
});

test("markdownToPlain keeps link text and drops the URL", () => {
  assert.equal(markdownToPlain("see [the docs](https://example.com)"), "see the docs");
});

test("markdownToPlain unwraps inline code", () => {
  assert.equal(markdownToPlain("run `npm test` now"), "run npm test now");
});
