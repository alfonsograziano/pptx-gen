# 0003 — The workspace is never the install

**Status:** Accepted

## Context

Two kinds of thing live in this system and they have opposite lifecycles. The **engine** is code that should be updated often, from a shared upstream. A **workspace** is the user's brand, template library, assets and decks — data that must never be overwritten by an update, and that may be one of several (a brand per client).

An earlier arrangement treated the repo root as a workspace. That folded the two together: an engine update sat in the same tree as someone's client decks, a path resolved relative to the install broke the moment the tool was run from somewhere else, and `git pull` and "my deck" became the same operation.

## Decision

A workspace is any folder containing a `pptx-gen.config.yml`. It is always a separate folder from the install, and the engine never writes into itself.

`resolveWorkspaceSync` in `src/workspace.ts` is the only way to find one. Precedence: the `--workspace` flag, then `$PPTX_GEN_WORKSPACE`, then the nearest config walking up — from the entry script's directory first, then the cwd — then `~/.pptx-gen`. A relative path inside a config resolves against the config file's own directory, never against `process.cwd()`; removing that cwd-dependence is what the feature exists for.

**Never resolve a path relative to the install, and never write deck files into it.** Go through `resolveWorkspaceSync` or `pptx-gen workspace --json`.

## Consequences

- Updating the engine is `git pull` in the clone. Nothing a user owns is in that tree.
- One install serves any number of workspaces, which is how a brand per client works.
- Every path-consuming code path needs a workspace, which is why `Presentation` resolves one lazily — pass an explicit `templateLibrary` and it needs none at all, which is how the test suite runs against `test/fixtures/workspace/`.
- A missing workspace has to fail well, since it is the first thing a new user hits. `WorkspaceNotFoundError` lists every directory searched and the three ways to fix it.
- Skills and documentation cannot hardcode paths; they ask `pptx-gen workspace --json` for `templates`, `projects`, `designDoc`, `customize` and the rest.
- `examples/workspace/` lives inside the repo, which looks like an exception and is not: it is a real workspace kept here as documentation, pointing its `templates` and `assets` back at the install so the repo carries no duplicated binaries.
