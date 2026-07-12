---
name: ingest-slide-templates
description: Ingest one or more slides from a PowerPoint deck into the pptx-gen template library and document each one. Use when the user gives a source .pptx file (often a single slide or a deck downloaded from Google Slides) plus a list of slide numbers and a template name, and asks to import, ingest, or add slides as reusable templates in templates/. Runs the ingest CLI for each slide, then runs the draft-slide-template-description workflow on each imported template so it ships with a real description instead of a TODO stub.
---

# Ingest slide templates

## Purpose

Turn slides from a real deck into reusable templates in `templates/`, fully
documented and ready to pick when building a deck.

This skill chains two steps:

1. **Ingest** each requested slide with the CLI (`ingest --slide`). This clones
   the slide, extracts its text fields and fonts, renders a screenshot (if
   LibreOffice is installed), and writes a `description.md` TODO stub.
2. **Describe** each imported template by running the
   `draft-slide-template-description` workflow, which replaces that stub with a
   short, accurate description.

The output is one finished template folder per slide, with `template.pptx`,
`template.yml`, `fields.yml`, an optional screenshot, and a real
`description.md`.

## Fixed paths

- Ingest CLI: run from the repo root with `npm run cli -- ingest ...`
- Template library: `templates/<template-name>/`

Use Node 20 or newer.

## Required inputs (stop if missing)

You need all three before doing anything. If any is missing, stop and ask.

1. **Source `.pptx` file path.** A local file. For Google Slides, the user
   exports it first (File > Download > Microsoft PowerPoint .pptx) and gives you
   the path. We do not download from cloud drives inside this skill: real decks
   are large, and a manual export is reliable.
2. **Slide numbers to ingest.** A list, for example `3, 5, 16`, or a single one.
3. **Template name(s).** What each imported slide should be called.
   - One slide: use the name as given (for example `two-column-comparison`).
   - Several slides: ask for one name per slide. If the user gives a single base
     name, append `-slide-NN` per slide (zero-padded), matching `--split`.

## Steps

### 1. Confirm inputs and inspect the source

Check the file exists and count how many slides it really has:

```bash
unzip -l "<source.pptx>" | grep -cE 'ppt/slides/slide[0-9]+\.xml'
```

Reconcile the numbers the user gave against this count:

- **One slide in the file.** Common: the user exported just the slide they care
  about, so the file holds only it even though they call it "slide 16". Ingest it
  as position `1`, and say so.
- **Many slides, all requested numbers in range.** Treat the numbers as
  positions and ingest each.
- **A requested number is out of range** and the file has more than one slide.
  Do not guess. Stop and ask which position they mean.

### 2. Ingest each slide

For each (position, template name) pair, run from the repo root. First confirm
the target name is free (`ls templates/<name>`); if it exists, ask before
overwriting.

```bash
npm run cli -- ingest \
  --source "<source.pptx>" \
  --template "<template-name>" \
  --slide <position>
```

The CLI ingests one slide per call, so loop it once per slide. After each ingest,
check the folder was created. If a `screenshot-warnings.md` appeared, read it and
tell the user (a missing screenshot just means LibreOffice is not installed; the
template is still valid).

To split every slide of a multi-slide deck in one call, use `--split` instead of
`--slide`.

### 3. Describe each imported template

For every template you just imported, run the
`draft-slide-template-description` workflow on it. Do each as its own job; do not
blend observations across slides.

### 4. Report

Tell the user, per slide: the template name and folder path, the field and font
counts from the ingestion report, a one-line summary of the description you
wrote, and any screenshot warnings.

## Notes and traps

- **The slide number is not always the file position.** A single-slide export is
  the normal case when the user grabs one slide from a big deck.
- **Do not invent layout.** The screenshot (or, if absent, the fields) is the
  source of truth for the description, not the slide name.
- **Template names are folder names.** Keep them kebab-case and descriptive of
  intent (`two-tier-comparison-cards`), not of the source deck.
