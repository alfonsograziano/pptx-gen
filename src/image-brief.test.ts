import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import {
  brandPalette,
  composeImagePrompt,
  NO_IMAGE_STYLE,
  parseImageRequests,
  parseImageStyle,
  sizeFor,
  updateImageNotes
} from "./image-brief.js";

const BRIEF = `# Deck

## Audience and objective

Operators.

## Images

\`\`\`yaml
- id: hero-city
  slide: 1
  description: Aerial view of a European city at dawn
  variants: 3
  pick: 2
- id: texture
  description: Soft paper texture
  aspect: wide
  references: inputs/refs/paper.jpg
\`\`\`

## Claims and numbers

None.
`;

test("the ## Images yaml block becomes requests, with defaults filled in", () => {
  const [hero, texture] = parseImageRequests(BRIEF);
  assert.deepEqual(hero, {
    id: "hero-city",
    description: "Aerial view of a European city at dawn",
    slide: 1,
    variants: 3,
    pick: 2,
    aspect: "landscape",
    size: undefined,
    style: undefined,
    references: [],
    palette: undefined
  });
  assert.equal(texture.variants, 1);
  assert.equal(texture.pick, 1);
  assert.equal(texture.aspect, "wide");
  assert.deepEqual(texture.references, ["inputs/refs/paper.jpg"], "a single path is accepted as a list of one");
});

test("a brief with no images is zero requests, not an error", () => {
  assert.deepEqual(parseImageRequests("# Deck\n\n## Narrative arc\n\nTODO\n"), []);
  assert.deepEqual(parseImageRequests("# Deck\n\n## Images\n\nNone yet.\n"), [], "a section without a fence");
  assert.deepEqual(parseImageRequests("## Images\n\n```yaml\n```\n"), [], "an empty fence");
});

test("a fence inside an HTML comment is an example, not a request", () => {
  const brief = "## Images\n\n<!-- e.g.\n```yaml\n- id: hero\n  description: x\n```\n-->\n";
  assert.deepEqual(parseImageRequests(brief), []);
});

test("a fence under a later section is not read as images", () => {
  const brief = "## Images\n\nNone.\n\n## Appendix\n\n```yaml\n- id: stray\n  description: x\n```\n";
  assert.deepEqual(parseImageRequests(brief), []);
});

test("malformed requests are authoring errors that name the entry and the problem", () => {
  const block = (yaml: string) => `## Images\n\n\`\`\`yaml\n${yaml}\n\`\`\`\n`;
  const cases: [string, RegExp][] = [
    ["- id: hero\n  description: x\n  varients: 2", /unknown key "varients" — did you mean "variants"\?/],
    ["- id: hero\n  description: x\n- id: hero\n  description: y", /'hero' is declared twice/],
    ["- id: Hero City\n  description: x", /"id" must be kebab-case/],
    ["- id: hero", /"description" is required/],
    ["- id: hero\n  description: x\n  variants: 5", /"variants" must be a whole number between 1 and 4/],
    ["- id: hero\n  description: x\n  variants: 2\n  pick: 3", /"pick" is 3, but only 2 variant\(s\)/],
    ["- id: hero\n  description: x\n  aspect: panoramic", /"aspect" must be one of/],
    ["- id: hero\n  description: x\n  size: 1000x700", /multiples of 16/],
    ["- id: hero\n  description: x\n  size: 1024x1024\n  aspect: square", /"aspect" or "size", not both/],
    ["id: hero", /must be a list/],
    ["- id: [unclosed", /does not parse/]
  ];
  for (const [yaml, message] of cases) {
    assert.throws(() => parseImageRequests(block(yaml), "/deck/brief.md"), message, yaml);
  }
  assert.throws(
    () => parseImageRequests(block("- id: hero"), "/deck/brief.md"),
    (error: Error) => error.message.startsWith("/deck/brief.md: image 'hero': "),
    "errors lead with the file and the entry"
  );
});

test("customize.md's image section yields guidance, example paths and the palette switch", () => {
  const customize = `# Customizations

## Rules

- Always ship on Fridays.

## Image generation

<!-- instructions for people, ignored -->
Editorial photography, soft natural light.
Muted tones.

Examples:
- assets/style/hero.jpg
- \`assets/style/two words.png\`
- ~/moodboard/third.webp — the best one

Palette: off
`;
  const style = parseImageStyle(customize, "/ws");
  assert.equal(style.guidance, "Editorial photography, soft natural light. Muted tones.", "soft wraps are joined");
  assert.equal(
    parseImageStyle("## Image generation\n\nCalm.\n\n- warm light\n- no people\n", "/ws").guidance,
    "Calm.\n\n- warm light\n- no people",
    "while paragraphs and list items keep their breaks"
  );
  assert.deepEqual(style.examples.slice(0, 2), [
    path.resolve("/ws", "assets/style/hero.jpg"),
    path.resolve("/ws", "assets/style/two words.png")
  ]);
  assert.ok(style.examples[2].endsWith(path.join("moodboard", "third.webp")));
  assert.ok(!style.examples[2].startsWith("/ws"), "~ is the home directory, not the workspace");
  assert.equal(style.palette, false);
});

test("a customize.md with no image section, or only a commented one, is the empty style", () => {
  assert.deepEqual(parseImageStyle("# Customizations\n\n## Rules\n", "/ws"), NO_IMAGE_STYLE);
  assert.deepEqual(
    parseImageStyle("## Image generation\n\n<!-- Examples:\n- assets/a.jpg\nPalette: off -->\n", "/ws"),
    NO_IMAGE_STYLE
  );
});

test("the prompt is composed in a fixed order from the brief, the style and the brand", () => {
  const [hero] = parseImageRequests(BRIEF);
  const style = { guidance: "Editorial photography.", examples: [], palette: true };
  const palette = brandPalette({ ink: "12182b", accent: "#3B82F6", accent2: "8B5CF6", accent3: "3B82F6" });

  assert.equal(
    composeImagePrompt({ ...hero, style: "Low sun, long shadows." }, style, { palette, hasReferences: true }),
    [
      "Aerial view of a European city at dawn",
      "Low sun, long shadows.",
      "Style: Editorial photography.",
      "Brand colours: where it suits the scene, lean on #12182B (ink), #3B82F6 (accent), #8B5CF6 (accent2).",
      "The attached images are style references only: match their look, palette and mood, but do not copy their subjects or composition.",
      "Format: a landscape (3:2) image for a presentation slide. No text, lettering, captions, logos or watermarks."
    ].join("\n\n")
  );

  const plain = composeImagePrompt(hero, { ...style, palette: false }, { palette, hasReferences: false });
  assert.doesNotMatch(plain, /Brand colours/, "customize.md can turn the palette off");
  assert.doesNotMatch(plain, /style references/);
  const perImage = composeImagePrompt(
    { ...hero, palette: true },
    { ...style, palette: false },
    {
      palette,
      hasReferences: false
    }
  );
  assert.match(perImage, /Brand colours/, "and one image can turn it back on");
});

test("aspects map to sizes the models accept, and an explicit size wins", () => {
  const [hero] = parseImageRequests(BRIEF);
  assert.equal(sizeFor(hero), "1536x1024");
  assert.equal(sizeFor({ ...hero, aspect: "portrait" }), "1024x1536");
  assert.equal(sizeFor({ ...hero, aspect: "square" }), "1024x1024");
  assert.equal(sizeFor({ ...hero, aspect: "wide" }), "1536x864");
  assert.equal(sizeFor({ ...hero, size: "1920x1088" }), "1920x1088");
});

test("notes land right after the yaml fence, and leave the rest of the brief alone", () => {
  const noted = updateImageNotes(BRIEF, {
    processed: ["hero-city"],
    notes: new Map([["hero-city", "generation failed — OpenAI returned 400."]])
  });
  const fenceEnd = noted.indexOf("```\n\n<!-- pptx-gen:image-notes:start -->");
  assert.ok(fenceEnd > 0, "the block follows the closing fence");
  assert.match(noted, /> - `hero-city`: generation failed — OpenAI returned 400\./);
  assert.ok(noted.startsWith(BRIEF.slice(0, BRIEF.indexOf("```\n\n## Claims"))));
  assert.ok(noted.endsWith("\n\n## Claims and numbers\n\nNone.\n"));

  assert.equal(
    updateImageNotes(noted, {
      processed: ["hero-city"],
      notes: new Map([["hero-city", "generation failed — OpenAI returned 400."]])
    }),
    noted,
    "rewriting the same notes changes nothing"
  );
  assert.equal(
    updateImageNotes(noted, { processed: ["hero-city", "texture"], notes: new Map() }),
    BRIEF,
    "once every note is resolved the block is gone and the brief is byte-for-byte what it was"
  );
});

test("a partial run only rewrites the notes of the images it processed", () => {
  const both = updateImageNotes(BRIEF, {
    processed: ["hero-city", "texture"],
    notes: new Map([
      ["hero-city", "failed one way."],
      ["texture", "failed another."]
    ])
  });
  const onlyHero = updateImageNotes(both, { processed: ["hero-city"], notes: new Map() });
  assert.doesNotMatch(onlyHero, /`hero-city`:/);
  assert.match(onlyHero, /> - `texture`: failed another\./, "an image outside --only keeps its note");

  const general = updateImageNotes(onlyHero, { processed: [], notes: new Map(), general: "No key." });
  assert.match(general, /> - No key\.\n> - `texture`/, "the general note comes first");
  const cleared = updateImageNotes(general, { processed: [], notes: new Map() });
  assert.doesNotMatch(cleared, /No key/, "and is replaced on every run");
});

test("a brief with no ## Images fence is never annotated", () => {
  const brief = "# Deck\n\n## Narrative arc\n\nTODO\n";
  assert.equal(updateImageNotes(brief, { processed: [], notes: new Map([["x", "y"]]) }), brief);
});
