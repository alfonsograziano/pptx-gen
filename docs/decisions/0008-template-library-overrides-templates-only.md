# 0008 — `templateLibrary` overrides the template root and nothing else

**Status:** Accepted

## Context

`Presentation` takes four things from a workspace: the template root, the assets folder, the bundled icon set, and (through `src/design.ts`) the design. [Decision 0003](0003-the-workspace-is-never-the-install.md) says the engine must never resolve a path relative to the install, and it also says a bare checkout with no workspace anywhere has to keep working — that is how the test suite runs and how a first-time user gets a deck out before creating anything.

The constructor reconciled those two by treating `templateLibrary` as a signal that no workspace was wanted: pass one, and the workspace lookup was skipped entirely. Only the template root reads as "explicit" in that reading. `assetsDir` then fell through to the install's own `assets/`, which holds the bundled Lucide icons and nobody's logos.

So a deck that named its template library got its logos looked up in the wrong folder. `resolveLogo` returns `undefined` for a file it cannot find, and `addLogo` / `addFooter` / `addWordmark` return early on `undefined`, so the deck built, validated, reported no warning, and shipped with the brand mark missing from every slide. Seventeen build scripts in one workspace did exactly that, and the failure was found by looking at the slides.

## Decision

`templateLibrary` moves the template root. It does not mean "no workspace".

`Presentation` resolves a workspace in every case, with `tryResolveWorkspaceSync`, and uses it for any option the caller left unset. Each option keeps its own fallback, so a checkout with no workspace still builds: the template root falls back to `resolveWorkspaceSync()` — called strictly, so a caller who supplied neither gets `WorkspaceNotFoundError` and its list of places searched — and assets fall back to the install's `assets/`.

**Every workspace-derived default is resolved independently. One explicit option never switches off another.**

And a logo that is asked for and not found is a `logo-not-found` warning in the build report, naming the file and the folder searched. `AssetResolver.missingLogos()` records the misses; `Presentation` drains them into `report.warnings` at the end of a build. Logos stay optional and a miss never fails the build, per [decision 0006](0006-external-renderers-are-optional.md) — but silence was the reason this cost seventeen decks instead of one.

## Consequences

- A deck script can point at a shared template library and keep its own brand. That combination was the whole reason `templateLibrary` existed as a public option, and it did not work.
- Resolving a workspace is no longer free, since it happens on every `new Presentation()`. It is a synchronous walk up the directory tree, already paid by `src/design.ts` at import time, so the cost is a second stat walk and not worth caching yet.
- A caller who genuinely wants the install's assets and no workspace has to say so, by passing `assetsDir`. There is no way to ask for "no workspace at all" any more, and there should not be: every path still has an install-relative fallback, so nothing requires one.
- A workspace with no logo files now warns on every build. That is the correct reading of decision 0006's rule — an optional thing that was skipped is named, not hidden — and the message says which folder to drop the PNGs into.
- Anything else `Presentation` learns to take from a workspace must follow the same shape: its own option, its own fallback, resolved on its own. Reintroducing a flag that gates several defaults at once brings this class of bug straight back.
