# 0010 — Generated images are inputs, made before the build

**Status:** Accepted

## Context

Some decks need photographs or illustrations: a cover picture, a scene, a texture. Nothing else in pptx-gen can make those. Native shapes draw diagrams, HTML figures draw interfaces and charts, and until now the only answer for a picture was a grey placeholder for someone to fill in later. An image model can fill that gap. But calling one is slow, costs money, needs a secret, and gives different pixels every time. Every other part of a build is offline and deterministic, and build scripts are written to stay that way (`skills/pptx-deck/SKILL.md`: "no LLM calls, network calls ... during render").

## Decision

Generation is its own command, `pptx-gen images`, and it runs **before** the build, never inside it.

- **Input.** The deck's `brief.md` declares each image in a fenced yaml block under `## Images`. The workspace's `customize.md` can add an `## Image generation` section with style guidance and example images. `image-brief.ts` composes one prompt per image from those, deterministically.
- **Output.** `image-gen.ts` calls the OpenAI Images API with plain `fetch`, no SDK, and writes the variants into the deck's `inputs/` as `<id>-<n>.jpg`. `inputs/images.lock.json` records a hash of everything that shapes each result, so an unchanged image is never generated twice. A file in `inputs/` that the lock does not own was put there by a person, and is never overwritten without `--force`.
- **Placement.** The build places an image by id: `helpers.addImage(slide, { image })` on custom slides, and the `addImage` / `replaceImage` overrides on cloned ones. It reads only what is on disk (`images.ts`). It never looks for a key, so the same inputs build the same deck on any machine.
- **Optional.** With no `OPENAI_API_KEY`, in the environment or the workspace `.env`, the command generates nothing, exits 0, and writes a note into the brief. A declared image with no file becomes the same captioned placeholder a figure falls back to, plus an `image-missing` warning. This is ADR 0006's rule applied to a network service instead of a binary.
- **Failures are written down.** They go into a managed notes block in the brief, next to what was asked for, and are removed once resolved. Only a malformed brief or `customize.md` is an error.

## Consequences

- Build scripts stay deterministic and runnable offline. A deck can be reviewed, rebuilt and diffed without the key, and CI never needs a secret.
- **ADR 0002 holds.** A generated image is content, like a photo the user supplies. It is never a diagram, a chart or text: every prompt forbids lettering, and the skill routes those visuals to native shapes and figures first.
- **ADR 0001 has one exception.** `replaceImage { image }` with no file yet swaps the template's picture for a placeholder rather than keeping the template's sample image. The sample is almost always stock art that would ship unnoticed, whereas a captioned grey box cannot. The placeholder keeps the picture's shape id and name, so later overrides still match. When the picture's box cannot be read, the original is kept and a warning says so.
- A `path` is still a claim that the file exists, so a missing one still fails the build. Only an `image` id is allowed to be missing.
- The engine now reads secrets. `init` ignores `.env` in new workspaces, and `doctor` flags (and `--fix` repairs) an older workspace whose `.env` is not ignored. The key is never printed, and a key fragment quoted back by the API is redacted before it can reach the brief.
- References are sent through the edits endpoint, which uses them as content as much as style. Scenery from an example can reappear in the result, whatever the prompt says. The skill says so.
- Changing the model or its API surface touches `image-gen.ts` only. The spec format, the lock and placement do not depend on the provider.
