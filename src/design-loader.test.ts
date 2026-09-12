import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { C, FONTS, LAYOUT, LOGO_FILES, applyDesign, currentDesign } from "./design.js";
import { parseDesignFile, readDesignFileSync, serializeDesign } from "./design-loader.js";

const DESIGN = "/ws/design.yml";

/** Restore the live design after a test mutates it. */
function withRestoredDesign(t: { after: (fn: () => unknown) => void }): void {
  const snapshot = currentDesign();
  t.after(() => applyDesign(snapshot));
}

test("an omitted token keeps the engine default", () => {
  const patch = parseDesignFile("colors:\n  accent: E8452C\n", DESIGN);
  assert.deepEqual(patch.colors, { accent: "E8452C" });
  assert.equal(patch.fonts, undefined);
  assert.equal(patch.layout, undefined);
});

test("an empty design file is a valid no-op patch", () => {
  assert.deepEqual(parseDesignFile("", DESIGN), {});
  assert.deepEqual(parseDesignFile("# just a comment\n", DESIGN), {});
});

test("a leading # is stripped and hex is normalised to upper case", () => {
  const patch = parseDesignFile('colors:\n  ink: "#0a0b0c"\n', DESIGN);
  assert.deepEqual(patch.colors, { ink: "0A0B0C" });
});

test("digits-only colours survive YAML's number coercion", () => {
  // YAML's core schema would read these as numbers, lossily: 000000 -> 0,
  // 001122 -> 1122, 0x1122 -> 4386. The literal text must win.
  assert.deepEqual(parseDesignFile("colors:\n  ink: 112233\n", DESIGN).colors, { ink: "112233" });
  assert.deepEqual(parseDesignFile("colors:\n  ink: 000000\n", DESIGN).colors, { ink: "000000" });
  assert.deepEqual(parseDesignFile("colors:\n  ink: 001122\n", DESIGN).colors, { ink: "001122" });
  assert.deepEqual(parseDesignFile("colors:\n  ink: 0E1234\n", DESIGN).colors, { ink: "0E1234" });
  // Quoting still works, for anyone who reaches for it out of habit.
  assert.deepEqual(parseDesignFile('colors:\n  ink: "000000"\n', DESIGN).colors, { ink: "000000" });
});

test("a value that is not a hex colour is still rejected", () => {
  assert.throws(() => parseDesignFile("colors:\n  ink: 0x1122\n", DESIGN), /must be a 6-digit hex colour/);
  assert.throws(() => parseDesignFile("colors:\n  ink: 12345\n", DESIGN), /must be a 6-digit hex colour such as "3B82F6", got "12345"\./);
});

test("a bad hex colour names the offending key", () => {
  assert.throws(
    () => parseDesignFile("colors:\n  accent: reddish\n", DESIGN),
    /^DesignFileError: \/ws\/design\.yml: colors\.accent must be a 6-digit hex colour such as "3B82F6", got "reddish"\.$/
  );
});

test("the raised-surface and on-dark tokens round-trip", () => {
  // A brand with off-white paper needs a card tone distinct from the page, and
  // an accent that clears contrast on dark. Without these the palette forces
  // one accent to work on both grounds.
  const patch = parseDesignFile(
    ["colors:", "  paperSoft: FFFFFF", "  inkSoft: 1F1F1F", "  accentOnDark: EF4A3F"].join("\n"),
    DESIGN
  );
  assert.deepEqual(patch.colors, { paperSoft: "FFFFFF", inkSoft: "1F1F1F", accentOnDark: "EF4A3F" });
});

test("an unknown colour key lists the valid ones", () => {
  assert.throws(
    () => parseDesignFile("colors:\n  primary: 3B82F6\n", DESIGN),
    /unknown key "primary" in colors\. Valid keys: ink, accent, white, accent2/
  );
});

test("an unknown section is rejected", () => {
  assert.throws(() => parseDesignFile("colours:\n  ink: 000000\n", DESIGN), /unknown section "colours"\. Valid sections: colors, fonts, layout, logos\./);
});

test("layout values must be positive numbers", () => {
  assert.throws(() => parseDesignFile("layout:\n  LM: 0\n", DESIGN), /layout\.LM must be a positive number of inches, got a number\./);
  assert.throws(() => parseDesignFile('layout:\n  LM: "0.75"\n', DESIGN), /layout\.LM must be a positive number of inches, got "0\.75"\./);
  assert.deepEqual(parseDesignFile("layout:\n  LM: 1.5\n", DESIGN).layout, { LM: 1.5 });
});

test("fonts and logos accept names and reject empties", () => {
  assert.deepEqual(parseDesignFile("fonts:\n  sans: Roboto\n", DESIGN).fonts, { sans: "Roboto" });
  assert.deepEqual(parseDesignFile("logos:\n  markDark: acme.png\n", DESIGN).logos, { markDark: "acme.png" });
  assert.throws(() => parseDesignFile('fonts:\n  sans: ""\n', DESIGN), /fonts\.sans must be a font family name/);
});

test("applyDesign is visible through a reference captured before the call", (t) => {
  withRestoredDesign(t);

  // This is the property the whole scheme rests on: a project's custom.ts does
  // `const { LM } = LAYOUT` and holds `C` at module scope.
  const colorsRef = C;
  const layoutRef = LAYOUT;

  applyDesign({ colors: { accent: "E8452C" }, layout: { LM: 1.5 } });

  assert.equal(colorsRef.accent, "E8452C");
  assert.equal(layoutRef.LM, 1.5);
  assert.equal(C.accent, "E8452C");
});

test("applyDesign leaves untouched tokens at their defaults", (t) => {
  withRestoredDesign(t);
  const inkBefore = C.ink;
  const serifBefore = FONTS.serif;

  applyDesign({ colors: { accent: "E8452C" } });

  assert.equal(C.ink, inkBefore);
  assert.equal(FONTS.serif, serifBefore);
  assert.equal(LOGO_FILES.markDark, "logo-mark-dark.png");
});

test("serializeDesign round-trips through parseDesignFile", (t) => {
  withRestoredDesign(t);
  applyDesign({ colors: { accent: "E8452C" }, fonts: { sans: "Roboto" }, layout: { LM: 1.25 } });

  const patch = parseDesignFile(serializeDesign(currentDesign()), DESIGN);

  assert.equal(patch.colors?.accent, "E8452C");
  assert.equal(patch.fonts?.sans, "Roboto");
  assert.equal(patch.layout?.LM, 1.25);
  // A serialized file is complete, so it carries every token.
  assert.equal(patch.colors?.ink, C.ink);
});

test("a missing design file means defaults, not an error", async (t) => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "pptx-design-test-"));
  t.after(() => rm(dir, { recursive: true, force: true }));

  assert.equal(readDesignFileSync(path.join(dir, "design.yml")), undefined);

  await writeFile(path.join(dir, "design.yml"), "colors:\n  accent: 112233\n", "utf8");
  assert.deepEqual(readDesignFileSync(path.join(dir, "design.yml"))?.colors, { accent: "112233" });
});
