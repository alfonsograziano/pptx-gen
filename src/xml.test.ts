import { test } from "node:test";
import assert from "node:assert/strict";
import { asArray, escapeXml, unescapeXml } from "./xml.js";

test("escapeXml escapes the XML metacharacters", () => {
  assert.equal(escapeXml("a < b & c > d"), "a &lt; b &amp; c &gt; d");
});

test("unescapeXml reverses the common entities", () => {
  assert.equal(unescapeXml("a &lt; b &amp; c &gt; d"), "a < b & c > d");
});

test("escape then unescape round-trips", () => {
  const original = "Tom & Jerry <tag> \"quote\"";
  assert.equal(unescapeXml(escapeXml(original)), original);
});

test("asArray normalises undefined, scalar, and array inputs", () => {
  assert.deepEqual(asArray(undefined), []);
  assert.deepEqual(asArray("x"), ["x"]);
  assert.deepEqual(asArray(["x", "y"]), ["x", "y"]);
});
