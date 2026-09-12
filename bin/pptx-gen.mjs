#!/usr/bin/env node
// The `pptx-gen` command.
//
// The engine is shipped as TypeScript and run through tsx rather than being
// compiled, so this registers the tsx loader and then hands over to the CLI.
// It exists because package.json `bin` cannot point at a .ts file — node
// refuses to execute one directly.
import { register } from "tsx/esm/api";

register();

await import(new URL("../src/cli.ts", import.meta.url).href);
