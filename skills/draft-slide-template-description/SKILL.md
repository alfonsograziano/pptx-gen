---
name: draft-slide-template-description
description: Draft the description.md for a slide template in templates/ by looking at its rendered screenshot (or its fields when no screenshot exists). Use whenever the user wants to describe, document, or write up what a slide template looks like and when to use it, fill in a template's TODO description stub, or asks "what is this slide for" about a template in the library. Triggers on a template name, or a request to document one or more slide templates. If no name is given, ask for it; if several are given, do each one.
---

# Draft slide template description

## Purpose

Every template folder under `templates/` ships a `description.md` that tells a
future deck builder what the slide looks like and when to reach for it. Freshly
ingested templates carry only a TODO stub. This skill replaces that stub with a
short, accurate description.

The point is selection speed: when someone builds a deck, they skim these files
to pick the right slide. So the description must be honest about the real layout
and useful for "should I use this one or not", not a generic blurb.

The screenshot is the source of truth. Describe what is really on the slide, not
what the name suggests. If there is no screenshot (LibreOffice was not installed
at ingest time), fall back to the fields and geometry, and say the description
was drafted from fields rather than a render.

## Inputs

- A **template name** is a folder name under `templates/`.
- If the request gives **no name**, list a few available folders and ask which to
  document.
- If it gives **several**, do each independently. Do not blend observations
  across slides.

## Steps

For each template:

1. **Resolve the folder** `templates/<name>/`. If it does not exist, list the
   available folders and ask the user to confirm the name.

2. **Look at the screenshot** at `screenshots/slide-01.png` if present. Read it
   visually: where the title sits, how many text blocks there are, columns or
   repeated pillars, icons, images, quotes, and what is emphasised (size, colour,
   the accent).

3. **Skim `fields.yml` and `template.yml`.** The first field's `originalText`
   often names the slide's intent. The number of fields and their geometry show
   how much content the slide holds. When there is no screenshot, these are your
   primary evidence.

4. **Write `description.md`** using the format below. Overwrite the TODO stub.

## Output format

Keep it short. Use this shape, and drop any section that would only state the
obvious:

```
# <Human-readable slide name>

<One sentence: what this slide is for.>

## Layout
<2-4 lines or bullets on the visible structure: title placement, number of text
blocks, columns/pillars, icons, images, emphasis.>

## When to use
<Bullets: the concrete situations this slide fits.>

## When not to use
<Optional. Only when there is a real trap: name the better alternative.>

## Fields
<Bullets: each editable field id and what it holds. This is what a builder types
into `variables`.>
```

Guidance:

- **Ground every claim in evidence.** If you cannot see it in the screenshot or
  confirm it in the fields, do not assert it.
- **Give a sense of capacity.** "Up to about four bullets", "two short columns":
  builders need to know how much copy fits before it overflows.
- **Be concrete about intent.** "Section opener", "side-by-side comparison" beats
  "a flexible content slide".
- **List the fields.** Deck builders fill fields by id, so an accurate field list
  is the most useful part when there is no screenshot.
