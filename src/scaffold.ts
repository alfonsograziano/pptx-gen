// Scaffolding a deck project inside a workspace.
//
// The generated build.ts imports the engine by its bare specifier and derives
// its own projectDir from import.meta.url, so the folder can be moved or
// copied to another workspace and still build.
import path from "node:path";
import { existsSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import type { Workspace } from "./workspace.js";

export type ScaffoldOptions = {
  workspace: Workspace;
  deckId: string;
  title?: string;
  custom?: boolean;
  force?: boolean;
};

export type ScaffoldResult = {
  projectDir: string;
  created: string[];
};

export async function scaffoldProject(options: ScaffoldOptions): Promise<ScaffoldResult> {
  const deckId = options.deckId.trim();
  if (!/^[a-z0-9][a-z0-9-]*$/.test(deckId)) {
    throw new Error(`Deck id "${options.deckId}" must be kebab-case: lower-case letters, digits and dashes.`);
  }

  const projectDir = path.join(options.workspace.projectsDir, deckId);
  if (existsSync(projectDir) && !options.force) {
    throw new Error(`${projectDir} already exists. Pass --force to add missing files to it.`);
  }

  const title = options.title ?? humanTitle(deckId);
  const created: string[] = [];

  await mkdir(path.join(projectDir, "inputs"), { recursive: true });
  await mkdir(path.join(projectDir, "output"), { recursive: true });

  await write(path.join(projectDir, "build.ts"), buildScript(title, options.custom === true), created);
  await write(path.join(projectDir, "brief.md"), briefStub(title), created);
  if (options.custom) await write(path.join(projectDir, "custom.ts"), customStub(), created);

  return { projectDir, created };
}

function buildScript(title: string, custom: boolean): string {
  const customImport = custom ? '\nimport { openingSlide } from "./custom.js";' : "";
  const customUse = custom
    ? `\ndeck.addCustomSlide(openingSlide({ title: ${JSON.stringify(title)} }));\n`
    : "";

  return `// Build script for "${title}".
//
//   pptx-gen build --script <this file>
//
// Deterministic on purpose: no network calls, no changing dates, no random
// values. The same inputs must always produce the same deck.
import path from "node:path";
import { fileURLToPath } from "node:url";
import { Presentation, md } from "pptx-gen";${customImport}

const deck = new Presentation({
  title: ${JSON.stringify(title)},
  // The deck's own folder, so output paths below are relative to it wherever
  // this script is run from. Templates and assets come from the workspace.
  projectDir: path.dirname(fileURLToPath(import.meta.url))
});
${customUse}
deck.addSlideFromTemplate({
  templateName: "title-cover",
  variables: {
    "overline-label": "Draft",
    "your-presentation-title-goes-here": ${JSON.stringify(title)},
    "a-short-subtitle-that-sets-up-the-st": md("Replace this with the real subtitle.")
  }
});

await deck.render({
  output: "output/deck.pptx",
  report: "output/report.md",
  screenshots: "output/screenshots"
});
`;
}

function briefStub(title: string): string {
  return `# ${title}

## Audience and objective

TODO

## Narrative arc

TODO

## Must-have slides

TODO

## Claims and numbers

List every figure with its source. Do not invent facts or inflate numbers.

| Claim | Number | Source |
| --- | --- | --- |
| TODO | TODO | TODO |
`;
}

function customStub(): string {
  return `// Slides designed from scratch for this deck.
//
// Read the engine's custom-template-instructions.md first (its path is in
// \`pptx-gen workspace --json\`). Build from native shapes, text and vector
// icons so everything stays editable in PowerPoint and Google Slides.
import { CustomSlide, C, LAYOUT } from "pptx-gen";

// Reading the design at top level is safe: the workspace design.yml is applied
// before this module is evaluated.
const { LM, CW } = LAYOUT;

export function openingSlide(input: { title: string }): CustomSlide {
  return new CustomSlide({
    name: "opening",
    background: "light",
    draw({ slide, helpers, pageNum }) {
      helpers.addHeader(slide, input.title);
      helpers.addTextBlock(
        slide,
        [{ text: "Replace this with the real opening statement." }],
        { x: LM, y: 1.6, w: CW, h: 1.2 },
        { fontSize: 28, color: C.ink }
      );
      helpers.addFooter(slide, pageNum);
    }
  });
}
`;
}

function humanTitle(deckId: string): string {
  return deckId.replace(/[-_]+/g, " ").replace(/^./, (letter) => letter.toUpperCase());
}

async function write(target: string, contents: string, created: string[]): Promise<void> {
  if (existsSync(target)) return;
  await writeFile(target, contents, "utf8");
  created.push(target);
}
