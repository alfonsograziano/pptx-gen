# AGENTS.md

pptx-gen is a TypeScript engine that compiles real, editable `.pptx` decks two ways: by cloning a slide's XML out of a template `.pptx` and filling its text fields, or by drawing a slide from native shapes. It runs `.ts` directly through `tsx` — there is no build step and no emitted JavaScript.

## Commands

```bash
npm run check            # the whole gate: lint, typecheck, tests, dead-code scan (~40s)
npm test                 # 114 tests, node --test over src/**/*.test.ts (~20s)
npm run build            # typecheck only: tsc --noEmit
npm run lint             # biome check (lint + format check); npm run lint:fix writes the fixes
npm run knip             # dead code: unused files, exports and dependencies
npm run example          # rebuild both example decks in examples/workspace
npm run self-validate    # end-to-end: ingest a slide, build a deck, check the package
npx tsx src/cli.ts ...   # the CLI without linking: init, workspace, where, new, doctor, ingest, build, validate
```

Run `npm run check` before handing work back — there is no CI, so it is the whole gate. Formatting is Biome's, configured in `biome.json` (120 columns, double quotes, no trailing commas); do not hand-format.

## Layout

- `src/` — the engine. `ooxml.ts` (1,100 lines) is the XML surgery, `presentation.ts` the deck builder, `custom-slide-helpers.ts` the drawing helpers, `design.ts` + `design-loader.ts` the design system, `workspace.ts` the workspace resolver, `figure.ts` + `html-shot.ts` the headless-Chrome figures. `index.ts` is the public API — anything not exported there is internal.
- `src/*.test.ts` — tests live next to the code they test, not in `test/`.
- `test/fixtures/workspace/` — the workspace the test suite runs against (`PPTX_GEN_WORKSPACE` is set for `npm test`).
- `bin/pptx-gen.mjs` → `src/cli.ts` — the CLI entry point.
- `docs/` — how the engine is put together (`docs/architecture.md`: the module map and the build pipeline, step by step) and why (`docs/decisions/`: one record per decision that would be expensive to reverse, including the reasoning behind every rule below). Read the relevant record before changing something a rule protects; add one when you make a call that constrains later work.
- `starter/` — what `pptx-gen init` copies into a new workspace. `examples/workspace/` — a real workspace used as the worked example.
- `skills/` — the four agent skills shipped to *users* of pptx-gen. They are product, not instructions for this repo.
- `assets/icons/` — 1,900+ Lucide SVGs, vendored. Never hand-edit them.
- `.claude/skills/` — BMAD (bmm module, v6.12.0), 29 `bmad-*` planning and review skills, installed and tracked. `_bmad/` holds its config; `_bmad/config.toml` is the team's and is committed, `_bmad/config.user.toml` is personal and ignored. Its output lands in `_bmad-output/`, which is not yet tracked either way. Run `bmad-help` if you are not sure which one applies. Do not hand-edit anything under `_bmad/` outside `custom/` — the installer regenerates it.

## Rules

- **The engine and the workspace are separate.** A workspace (`pptx-gen.config.yml`, `design.yml`, `templates/`, `projects/`) lives outside this repo. Never write deck files into the install, and never resolve a path relative to the install — go through `resolveWorkspaceSync` / `pptx-gen workspace --json`.
- **Design values are data, not code.** They come from the workspace's `design.yml` layered over the defaults in `src/design.ts`. Adding a token means adding a default *and* letting a partial override fall back, so an older `design.yml` keeps working.
- **Slides stay native.** Everything the engine emits must be a shape, line or text run, because Google Slides cannot edit an embedded image. HTML figures are the one exception and the bar is high — see `custom-template-instructions.md` and `figure-instructions.md`, which are the full contracts for custom slides and figures.
- **Cloned slides keep their pixels.** Filling a field or applying an override edits the original slide's XML; it never redraws it. A change that makes a cloned slide look different from its source is a bug.
- **`any` is confined to the XML layer.** `XmlNode` in `src/xml.ts` is the one documented escape hatch, for the tree fast-xml-parser returns. New `any` anywhere else is a lint error, and should be.
- Generated output is gitignored (`output/screenshots/`, `output/figures/`, `*.pdf`), with one deliberate exception whitelisted in `.gitignore` so the repo browses without a browser installed. Do not commit render artifacts.
- LibreOffice (screenshots) and Chrome (figures) are optional. Anything that depends on them must degrade to a warning and still produce a deck.

## Ask first

`npm link`, `npm run install-fonts` (writes into the user's font directory), anything that writes outside this repo or into someone's workspace, and any push to `master`.
