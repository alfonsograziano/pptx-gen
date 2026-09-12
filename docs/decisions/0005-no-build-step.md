# 0005 — No build step: TypeScript runs directly

**Status:** Accepted

## Context

The primary user of this codebase is an AI coding agent driving a deck build, and the primary authoring surface is a `build.ts` the agent writes inside a workspace that is not this repo. A compile step would sit between every edit and every result: the agent edits, forgets to rebuild, runs stale JavaScript, and debugs the wrong code. A workspace script would need its own toolchain to import a package shipped as compiled output.

## Decision

The engine ships and runs as TypeScript. `tsx` strips the types at run time; nothing is emitted, and there is no `dist/`.

`package.json` points `main` and `exports` straight at `src/index.ts`. `bin/pptx-gen.mjs` launches `src/cli.ts` through tsx. `npm run build` is `tsc --noEmit` — a typecheck, not a build, and the name is kept only because it is where people look.

## Consequences

- Edit and run. Nothing is ever stale, and a workspace's `build.ts` imports `pptx-gen` and gets real types with no toolchain of its own.
- **Types are not checked at run time, so `tsc --noEmit` is the only thing that checks them at all.** This is why `npm run build` is inside `npm run check` and why the gate has to be run before handing work back.
- Module resolution is NodeNext, so relative imports are written with a `.js` extension even though the files on disk are `.ts`. This trips up everyone once.
- The `.ts` files are what ships. `package.json` lists `src` in `files`, and a consumer's own runner strips the types — so the published surface is source, and there is no emitted artifact whose contents could drift from it.
- Consumers need Node 20 or newer (`engines.node: ">=20"`, `.nvmrc` tracking the current LTS).
- A revert is live the moment it lands: there is no rebuild and no redeploy between a commit and the code that runs.
