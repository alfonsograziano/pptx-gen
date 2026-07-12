# Brief: "How pptx-gen works" explainer slide

**Audience.** Anyone landing on the repo who wants to understand, at a glance,
what pptx-gen does.

**Objective.** One editable slide that visually explains the whole pipeline:
you describe a deck to an AI agent, the engine builds each slide one of two ways
(clone a real slide, or design from scratch), and the output is a real,
editable `.pptx`.

**Why it ships in `examples/`.** It doubles as a live demo of the tool building
a diagram entirely from native shapes, cards, arrows, and vector icons — the
exact "design from scratch" mode it describes.

**Narrative (left to right):**

1. Your brief — describe the deck in plain language; an AI agent (driven by the
   skills) does the work.
2. Two build modes — *clone & fill* a real slide, or *design from scratch* with
   the design system. They mix freely in one deck.
3. The engine — assembles slides, embeds fonts, renders previews.
4. A real `.pptx` — fully editable, opens in PowerPoint, Keynote, and Google
   Slides; ships with screenshots and a build report.

**Constraints.** Native shapes only (editable in Google Slides). Sentence case.
Core palette plus `accent2`/`accent3` for the two modes.
