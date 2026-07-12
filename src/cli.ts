#!/usr/bin/env node
import { Command } from "commander";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { ingestTemplate } from "./ingest.js";
import { validatePackage } from "./ooxml.js";

const program = new Command();

program
  .name("nearform-pptx")
  .description("Nearform filesystem PPTX template generator")
  .version("0.1.0");

program
  .command("ingest")
  .requiredOption("--source <pptx>", "Source PPTX")
  .requiredOption("--template <name>", "Template name")
  .option("--slide <number>", "1-based slide number")
  .option("--split", "Import every slide as a separate flat template")
  .option("--template-root <dir>", "Template root", "templates")
  .action(async (options) => {
    const imported = await ingestTemplate({
      source: options.source,
      templateName: options.template,
      slide: options.slide === undefined ? undefined : Number(options.slide),
      split: options.split,
      templateRoot: options.templateRoot
    });
    console.log(`Imported ${imported.length} template(s) from ${options.source}`);
    for (const templateName of imported) console.log(`- ${templateName}`);
  });

program
  .command("build")
  .requiredOption("--script <file>", "Build script that exports or renders a Presentation")
  .action(async (options) => {
    const scriptPath = path.resolve(options.script);
    // The build script's deck.render() prints its own rich, timed summary,
    // so the CLI stays quiet here and lets that output speak for itself.
    await import(pathToFileURL(scriptPath).href);
  });

program
  .command("validate")
  .requiredOption("--pptx <file>", "PPTX file to validate")
  .action(async (options) => {
    await validatePackage(path.resolve(options.pptx));
    console.log(`Valid PPTX package: ${options.pptx}`);
  });

await program.parseAsync(process.argv);
