# pptx-gen documentation

This folder holds the reasoning behind the engine: how it is put together, and why it is put together that way. It is written for whoever — person or agent — has to change the code and needs to know which parts are load-bearing before they touch them.

| Document | What it answers |
| --- | --- |
| [architecture.md](architecture.md) | What the modules are, what happens during a build, and where each piece of the pipeline lives. |
| [decisions/](decisions/) | Why the system is shaped this way. One record per decision that would be expensive to reverse. |

## What belongs here, and what does not

These documents record **decisions and structure**. They are not a user manual and not an API reference, because those already exist and duplicating them guarantees one copy goes stale:

- **Using pptx-gen** — [`README.md`](../README.md) at the repo root: install, the CLI, the deck-authoring API, the skills.
- **Working in this repo** — [`AGENTS.md`](../AGENTS.md): the commands, the layout, and the rules an agent needs loaded before it edits anything.
- **Authoring a custom slide** — [`custom-template-instructions.md`](../custom-template-instructions.md), the full contract for slides drawn from native shapes.
- **Authoring a figure** — [`figure-instructions.md`](../figure-instructions.md), the full contract for HTML figures.
- **The design system as a brand owner sees it** — [`starter/design.md`](../starter/design.md).

If a fact lives in one of those, link to it from here rather than restating it.

## Adding a decision record

When you make a call that constrains later work — an output format, a dependency the code is shaped around, a boundary that must not be crossed — write it down in [`decisions/`](decisions/). The format and the bar for writing one are in [decisions/README.md](decisions/README.md).
