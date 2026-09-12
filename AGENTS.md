# AGENTS.md

pptx-gen is a TypeScript engine that compiles real, editable `.pptx` decks two ways: by cloning a slide's XML out of a template `.pptx` and filling its text fields, or by drawing a slide from native shapes. It runs `.ts` directly through `tsx` — there is no build step and no emitted JavaScript.

## Commands

```bash
npm test                 # 114 tests, node --test over src/**/*.test.ts (~20s)
npm run build            # typecheck only: tsc --noEmit
npm run example          # rebuild both example decks in examples/workspace
npm run self-validate    # end-to-end: ingest a slide, build a deck, check the package
npx tsx src/cli.ts ...   # the CLI without linking: init, workspace, where, new, doctor, ingest, build, validate
```

Run `npm run build && npm test` before handing work back. There is no linter, no formatter and no CI, so those two commands are the whole gate.

## Layout

- `src/` — the engine. `ooxml.ts` (1,100 lines) is the XML surgery, `presentation.ts` the deck builder, `custom-slide-helpers.ts` the drawing helpers, `design.ts` + `design-loader.ts` the design system, `workspace.ts` the workspace resolver, `figure.ts` + `html-shot.ts` the headless-Chrome figures. `index.ts` is the public API — anything not exported there is internal.
- `src/*.test.ts` — tests live next to the code they test, not in `test/`.
- `test/fixtures/workspace/` — the workspace the test suite runs against (`PPTX_GEN_WORKSPACE` is set for `npm test`).
- `bin/pptx-gen.mjs` → `src/cli.ts` — the CLI entry point.
- `starter/` — what `pptx-gen init` copies into a new workspace. `examples/workspace/` — a real workspace used as the worked example.
- `skills/` — the four agent skills shipped to *users* of pptx-gen. They are product, not instructions for this repo.
- `assets/icons/` — 1,900+ Lucide SVGs, vendored. Never hand-edit them.

## Rules

- **The engine and the workspace are separate.** A workspace (`pptx-gen.config.yml`, `design.yml`, `templates/`, `projects/`) lives outside this repo. Never write deck files into the install, and never resolve a path relative to the install — go through `resolveWorkspaceSync` / `pptx-gen workspace --json`.
- **Design values are data, not code.** They come from the workspace's `design.yml` layered over the defaults in `src/design.ts`. Adding a token means adding a default *and* letting a partial override fall back, so an older `design.yml` keeps working.
- **Slides stay native.** Everything the engine emits must be a shape, line or text run, because Google Slides cannot edit an embedded image. HTML figures are the one exception and the bar is high — see `custom-template-instructions.md` and `figure-instructions.md`, which are the full contracts for custom slides and figures.
- **Cloned slides keep their pixels.** Filling a field or applying an override edits the original slide's XML; it never redraws it. A change that makes a cloned slide look different from its source is a bug.
- Generated output is gitignored (`output/screenshots/`, `output/figures/`, `*.pdf`), with one deliberate exception whitelisted in `.gitignore` so the repo browses without a browser installed. Do not commit render artifacts.
- LibreOffice (screenshots) and Chrome (figures) are optional. Anything that depends on them must degrade to a warning and still produce a deck.

## Ask first

`npm link`, `npm run install-fonts` (writes into the user's font directory), anything that writes outside this repo or into someone's workspace, and any push to `master`.
