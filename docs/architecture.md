# Architecture

pptx-gen compiles a `.pptx` file two ways, and the whole design follows from the fact that both ways have to produce slides that stay editable in PowerPoint *and* Google Slides.

- **Clone and fill.** Copy a real slide's XML out of a template `.pptx` and edit only the parts asked for. The slide is never redrawn, so it keeps its fonts, shapes, arrows and images exactly. See [decision 0001](decisions/0001-clone-slide-xml-rather-than-redraw.md).
- **Design from scratch.** Draw a slide from native shapes, lines, text runs and vector icons, guided by the design system. See [decision 0002](decisions/0002-native-shapes-are-the-output-contract.md).

Both kinds end up in the same output deck, because both are reduced to the same thing before assembly: a single-slide `.pptx` package whose slide gets appended into the deck being built.

## The two entry points

| Entry point | What it is |
| --- | --- |
| `bin/pptx-gen.mjs` → `src/cli.ts` | The CLI. Commander, eight commands: `init`, `workspace`, `where`, `new`, `doctor`, `ingest`, `build`, `validate`. |
| `src/index.ts` | The public API. A deck's own `build.ts` imports `Presentation`, `CustomSlide`, the design tokens and the helpers from here. Anything not exported from `index.ts` is internal and may change. |

`pptx-gen build --script <file>` is a thin wrapper over the second: it resolves the workspace, applies its design, then imports the script, which does its own `deck.render(...)`.

## Startup: workspace, then design

Two things must be settled before any deck code runs, and the order matters.

**1. Find the workspace** (`src/workspace.ts`). A workspace is any folder holding a `pptx-gen.config.yml`; it owns the user's brand, template library, decks and assets, and is always a different folder from the install ([decision 0003](decisions/0003-the-workspace-is-never-the-install.md)). `resolveWorkspaceSync` picks one in strict precedence order:

1. the `--workspace` flag,
2. `$PPTX_GEN_WORKSPACE`,
3. the nearest `pptx-gen.config.yml` walking up — from the entry script's directory first, then the cwd,
4. `~/.pptx-gen`.

If none is found it throws `WorkspaceNotFoundError`, which lists every directory it searched and how to fix it. `tryResolveWorkspaceSync` is the non-throwing variant, used where a missing workspace must not be fatal. The config itself is parsed and validated by `src/workspace-config.ts`, which resolves every relative path against the config file rather than the cwd, expands a leading `~`, rejects unknown keys with a did-you-mean, and refuses a config version this install does not understand.

**2. Apply the design** (`src/design.ts`). Importing the module reads the workspace's `design.yml` and layers it over the built-in defaults, mutating `C`, `FONTS`, `LAYOUT` and `LOGO_FILES` **in place** ([decision 0004](decisions/0004-design-values-are-data-layered-over-defaults.md)). In-place mutation is what lets a deck script write `const { LM } = LAYOUT` at the top level and still see the brand's value.

The consequence is a real ordering rule, stated in `src/cli.ts`: when `--workspace` names a workspace other than the one found in the environment, its design must be applied *before* any project script is imported, and a design is never re-applied after that point. A script's top-level destructuring captures whatever the tokens hold at import time.

## The module map

| Layer | Modules | Job |
| --- | --- | --- |
| Package I/O | `pptx-package.ts` | A `.pptx` is a zip. Read, write, copy and remove entries through JSZip. Nothing above this layer touches the archive directly. |
| XML | `xml.ts`, `ooxml.ts` | `xml.ts` wraps fast-xml-parser and owns `XmlNode`, the one place `any` is allowed ([decision 0007](decisions/0007-any-is-confined-to-the-xml-layer.md)). `ooxml.ts` is the XML surgery: slicing, appending, field extraction, text filling, overrides, font merging, slide numbering, validation. |
| Templates | `templates.ts`, `ingest.ts` | Load a template directory; turn a source `.pptx` into one. |
| Deck building | `presentation.ts` | The orchestrator. Collects slides, runs the pipeline, writes the report. |
| Scratch slides | `custom-slide.ts`, `custom-slide-helpers.ts`, `svg-path.ts` | Draw a slide with pptxgenjs; convert SVG paths into native custom-geometry shapes. |
| Design | `design.ts`, `design-loader.ts`, `design-tokens.ts` | The live token values, reading a `design.yml`, and the token vocabulary. `design-tokens.ts` imports nothing, which keeps the other two from depending on each other. |
| Figures | `figure.ts`, `html-shot.ts` | Render author-written HTML to a PNG with headless Chrome. |
| Environment | `workspace.ts`, `workspace-config.ts`, `assets.ts`, `fonts.ts`, `install-fonts.ts`, `exec.ts`, `fs.ts` | Where things live and how to reach them. |
| Output | `render.ts`, `progress.ts` | LibreOffice screenshots; live terminal progress. |
| Setup | `init.ts`, `scaffold.ts` | `pptx-gen init` and `pptx-gen new`. |

## A build, step by step

`Presentation.render()` in `src/presentation.ts` is the whole pipeline. A deck with at least one template slide takes the main path; a deck of nothing but custom slides takes a shorter one (`renderAllCustom`) that skips the base package entirely.

1. **Pick a base package.** The first template slide's `template.pptx` is loaded and becomes the deck under construction. Its own slide entries are captured up front, so a later slide reusing the same template still sees the template's slides rather than the deck as it has grown.
2. **Append each slide in order.** Every template package is a single-slide slice sharing an identical support chain — layouts, masters, themes, fonts — with the base, which is what makes `appendSlideFromPackage` safe. A custom slide is first drawn to a throwaway single-slide `.pptx` by pptxgenjs, then appended through exactly the same call. From here on the pipeline cannot tell the two kinds apart.
3. **Fill and edit template slides.** `fillSlideText` writes the variables into the fields recorded in `fields.yml`; `convertTemplatePageNumber` swaps a field tagged `role: page-number` for a live PowerPoint slide-number field; `applyOverrides` applies the explicit edit operations — `delete`, `hide`, `move`, `resize`, `styleText`, `addText`, `addSvg`, `addIcon`, `addImage`, `replaceImage`, `addFigure`, `replaceFigure`.
4. **Carry fonts across.** `mergeEmbeddedFonts` copies any typeface a source slide embeds that the base lacks, so the delivered deck is self-contained.
5. **Assemble.** Unused figure files are pruned, `keepOnlySlides` trims the base down to the slides actually built, `applySlideNumbering` settles numbering now that the order is final, and the package is saved.
6. **Validate.** `validatePackage` re-opens the written file and checks it is a coherent package.
7. **Screenshot** (optional). LibreOffice converts the deck to PDF and `pdf-to-img` turns each page into a PNG. Absent LibreOffice, this is skipped with a warning and the build still succeeds ([decision 0006](decisions/0006-external-renderers-are-optional.md)). The copy that is shot has its slide-number fields flattened first, because LibreOffice resolves `slidenum` fields but ignores the deck's `firstSlideNum` offset and would otherwise show a number one too high on every slide of a deck with a cover. The delivered file keeps its live fields.
8. **Report.** A `BuildReport` records every slide in output order, variant-group membership, templates and custom slides used, figures and their status, and every warning. `output/report.md` is the human-readable form, and it is the artifact a reviewing agent reads to decide whether the deck is right.

Nothing in the pipeline throws on a degraded result. Anything recoverable becomes a `BuildWarning` with a `code`, and the deck is still produced — see the warning codes scattered through `ooxml.ts`, `figure.ts` and `render.ts`.

## Templates on disk

`pptx-gen ingest` (`src/ingest.ts`) turns one slide of a source deck into a template directory:

```
templates/<id>/
  template.pptx        the source deck sliced down to a single slide
  template.yml         TemplateMetadata: source, fonts, variables, tags, status
  fields.yml           FieldsFile: every text field, its id, shape id and original text
  description.md       a stub, filled in by the draft-slide-template-description skill
  ingestion-report.md  what was detected at import time
  screenshots/         rendered previews, when LibreOffice is available
```

Field ids are derived from the original text and are stable and unique, which is what lets a build address a field by name. A field detected as a page number is tagged `role: page-number` and deliberately left out of the metadata's `variables` list — the deck owns the page number, not the author.

## Figures

A figure is HTML the deck author writes, rasterized at build time and placed as a picture. It is the single deliberate exception to native-only output, and the bar for using one is high: it is for content with no native form at all — a product UI, a chart, a rendered document.

`figure.ts` assembles the final HTML (injecting the brand's colours and fonts), writes it next to the deck, and asks `html-shot.ts` to shoot it. Chrome is located by probing a list of well-known paths per platform, overridable with `$CHROME_PATH`, and the result is memoized so detection runs at most once.

Caching is content-addressed: the PNG's filename carries a hash of the generated HTML, the viewport, the scale, the transparency flag and the Chrome version — hence `alerts-console.4513df5752e2.png`. An unchanged figure is never re-shot, a changed one gets a new filename, and `prune()` deletes whatever the build did not use. Within one build, in-flight renders are keyed by output path, so the same figure on five slides is rasterized once.

Without a browser, each figure becomes a captioned grey placeholder using the figure's own `caption` — which is why the caption is specified as a sentence written for a reader, not a label.

## Assets

`assets.ts` resolves the three kinds of file a slide can refer to, and the differences between them are deliberate:

- **Icons** — the workspace's `assets/icons/` first, then the ~1,900 Lucide SVGs bundled with the engine. The workspace layer shadows and adds; copying the whole set into every workspace would bloat it and freeze it at init time.
- **Logos** — the workspace only, never a fallback. Falling back to the install could stamp the tool's placeholder mark onto a client's deck.
- **Project paths** — resolved against the deck's own folder.

A bare name like `rocket` is an icon lookup; anything containing a path separator is a project file.

## Testing

Tests live next to the code they test as `src/*.test.ts` and run on `node --test` with types stripped by tsx. `npm test` sets `PPTX_GEN_WORKSPACE=test/fixtures/workspace` so the suite always resolves a known workspace rather than whatever happens to be on the machine.

The suite works on real inputs — it loads the actual starter templates, builds real decks and re-opens the results. There is no mocking framework. The one seam that is injected is `ShotFn`, the figure rasterizer, so figure behaviour (including the no-browser path) can be tested without a browser installed.

`npm run check` is the whole gate: Biome, `tsc --noEmit`, the tests, then Knip for dead code. There is no CI, so it is the only gate.
