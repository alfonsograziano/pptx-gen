# Decision records

One file per decision that would be expensive to reverse. They exist so that a change made a year from now knows what it is trading away, and so that a rule in [`AGENTS.md`](../../AGENTS.md) has somewhere to point when someone asks *why*.

| # | Decision |
| --- | --- |
| [0001](0001-clone-slide-xml-rather-than-redraw.md) | A cloned slide is edited, never redrawn |
| [0002](0002-native-shapes-are-the-output-contract.md) | Native shapes are the output contract |
| [0003](0003-the-workspace-is-never-the-install.md) | The workspace is never the install |
| [0004](0004-design-values-are-data-layered-over-defaults.md) | Design values are data, layered over defaults |
| [0005](0005-no-build-step.md) | No build step: TypeScript runs directly |
| [0006](0006-external-renderers-are-optional.md) | External renderers are optional and degrade to warnings |
| [0007](0007-any-is-confined-to-the-xml-layer.md) | `any` is confined to the XML layer |

## Writing one

Copy the shape of an existing record: **Context** (the forces, not the history), **Decision** (what holds now, in the present tense), **Consequences** (what this costs, including what it forbids), and **Status**. Keep it to a page. Name real files and real functions — a record that cannot be checked against the code is a record nobody will trust in a year.

Number the next one sequentially. A superseded record is not deleted: set its status to `Superseded by NNNN` and leave the reasoning where it is, because the reason something was tried and abandoned is worth as much as the reason it was chosen.

Write a record when a choice constrains later work: an output format, a boundary that must not be crossed, a dependency the code is now shaped around. Do not write one for a choice that is cheap to change — that is what the code and its comments are for.
