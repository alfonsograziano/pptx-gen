#!/usr/bin/env node
import { Command } from "commander";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { ingestTemplate } from "./ingest.js";
import { validatePackage } from "./ooxml.js";
import { applyDesign, designOrigin } from "./design.js";
import { readDesignFileSync } from "./design-loader.js";
import { fixWorkspace, initWorkspace } from "./init.js";
import { scaffoldProject } from "./scaffold.js";
import { checkWorkspace, installDir, resolveWorkspaceSync, type Workspace } from "./workspace.js";

const program = new Command();

program
  .name("pptx-gen")
  .description("Filesystem-based PPTX template ingestion and deck generation")
  .version("0.1.0")
  .option(
    "-w, --workspace <dir>",
    "Workspace directory (default: nearest pptx-gen.config.yml, or $PPTX_GEN_WORKSPACE)"
  );

/**
 * Resolve the workspace for a command.
 *
 * NOTE ON ORDERING: importing `design.js` has already applied the design of the
 * workspace found in the environment. When --workspace names a DIFFERENT one,
 * its design must be applied here — before any project script is imported,
 * because a script's top-level `const { LM } = LAYOUT` captures whatever it
 * sees at that moment. Never re-apply a design after that point.
 */
function workspaceFor(command: Command, entryDir?: string): Workspace {
  const explicit = command.optsWithGlobals().workspace as string | undefined;
  return resolveWorkspaceSync({ explicit, entryDir });
}

program
  .command("init")
  .description("Create a workspace: config, brand, templates, projects, assets")
  .argument("[dir]", "Directory to create the workspace in", ".")
  .option("--force", "Refresh the scaffolding of an existing workspace")
  .option("--no-starter-templates", "Skip copying the bundled starter templates")
  .option("--no-link", "Skip the node_modules/pptx-gen engine link")
  .action(async (dir, options) => {
    const result = await initWorkspace({
      dir,
      force: options.force,
      starterTemplates: options.starterTemplates,
      link: options.link
    });

    console.log(`Workspace ready: ${result.workspace.root}`);
    for (const item of result.created) console.log(`  + ${path.relative(result.workspace.root, item) || "."}`);
    if (result.skipped.length > 0) {
      console.log(`  (${result.skipped.length} existing file(s) left alone)`);
    }
    console.log("");
    console.log("Next:");
    console.log(`  pptx-gen new my-first-deck --workspace ${result.workspace.root}`);
  });

program
  .command("workspace")
  .description("Show the resolved workspace paths")
  .option("--json", "Machine-readable output")
  .action(async (options, command: Command) => {
    const workspace = workspaceFor(command);
    const problems = await checkWorkspace(workspace);

    if (options.json) {
      const origin = designOrigin();
      console.log(
        JSON.stringify(
          {
            root: workspace.root,
            config: workspace.configPath,
            templates: workspace.templatesDir,
            projects: workspace.projectsDir,
            assets: workspace.assetsDir,
            icons: workspace.iconsDir,
            bundledIcons: workspace.bundledIconsDir,
            design: workspace.designPath,
            designDoc: workspace.designDocPath,
            customize: workspace.customizePath,
            designLoadedFrom: "file" in origin ? origin.file : null,
            customSlideGuide: path.join(workspace.installDir, "custom-template-instructions.md"),
            figureGuide: path.join(workspace.installDir, "figure-instructions.md"),
            install: workspace.installDir,
            engineSpecifier: "pptx-gen",
            source: workspace.source,
            ok: problems.length === 0,
            problems
          },
          null,
          2
        )
      );
      return;
    }

    console.log(`Workspace   ${workspace.root}  (${workspace.source})`);
    console.log(`Config      ${workspace.configPath}`);
    console.log(`Templates   ${workspace.templatesDir}`);
    console.log(`Projects    ${workspace.projectsDir}`);
    console.log(`Assets      ${workspace.assetsDir}`);
    console.log(`Design      ${workspace.designPath}`);
    console.log(`Design doc  ${workspace.designDocPath}`);
    console.log(`Customize   ${workspace.customizePath}`);
    console.log(`Install     ${workspace.installDir}`);
    if (problems.length > 0) {
      console.log("");
      console.log("Problems:");
      for (const problem of problems) console.log(`  - ${problem.message}`);
      console.log("");
      console.log("Run `pptx-gen doctor --fix` to repair what can be repaired.");
    }
  });

program
  .command("new")
  .description("Scaffold a deck project in the workspace")
  .argument("<deck-id>", "Kebab-case folder name, e.g. q4-board-review")
  .option("--title <title>", "Deck title")
  .option("--custom", "Also scaffold a custom.ts for slides designed from scratch")
  .option("--force", "Add missing files to an existing project folder")
  .action(async (deckId, options, command: Command) => {
    const workspace = workspaceFor(command);
    const result = await scaffoldProject({
      workspace,
      deckId,
      title: options.title,
      custom: options.custom,
      force: options.force
    });

    console.log(`Project ready: ${result.projectDir}`);
    for (const item of result.created) console.log(`  + ${path.relative(result.projectDir, item)}`);
    console.log("");
    console.log("Next:");
    console.log(`  pptx-gen build --script ${path.join(result.projectDir, "build.ts")}`);
  });

program
  .command("doctor")
  .description("Check the workspace, and optionally repair it")
  .option("--fix", "Repair what can be repaired")
  .action(async (options, command: Command) => {
    const workspace = workspaceFor(command);
    const problems = await checkWorkspace(workspace);

    if (problems.length === 0) {
      console.log(`Workspace is healthy: ${workspace.root}`);
      return;
    }

    console.log(`Workspace: ${workspace.root}`);
    for (const problem of problems) console.log(`  - ${problem.message}`);

    if (!options.fix) {
      console.log("");
      console.log("Run `pptx-gen doctor --fix` to repair these.");
      process.exitCode = 1;
      return;
    }

    const fixed = await fixWorkspace(workspace, problems);
    console.log("");
    for (const item of fixed) console.log(`  ✓ ${item}`);

    const remaining = await checkWorkspace(workspace);
    if (remaining.length > 0) {
      console.log("");
      console.log("Still broken:");
      for (const problem of remaining) console.log(`  - ${problem.message}`);
      process.exitCode = 1;
    }
  });

program
  .command("ingest")
  .requiredOption("--source <pptx>", "Source PPTX")
  .requiredOption("--template <name>", "Template name")
  .option("--slide <number>", "1-based slide number")
  .option("--split", "Import every slide as a separate flat template")
  .option("--template-root <dir>", "Template root (default: the workspace's templates/)")
  .action(async (options, command: Command) => {
    const templateRoot = options.templateRoot ?? workspaceFor(command).templatesDir;
    const imported = await ingestTemplate({
      source: options.source,
      templateName: options.template,
      slide: options.slide === undefined ? undefined : Number(options.slide),
      split: options.split,
      templateRoot
    });
    console.log(`Imported ${imported.length} template(s) from ${options.source}`);
    for (const templateName of imported) console.log(`- ${templateName}`);
  });

program
  .command("build")
  .requiredOption("--script <file>", "Build script that exports or renders a Presentation")
  .action(async (options, command: Command) => {
    const scriptPath = path.resolve(options.script);
    // Discover from the SCRIPT's location, not this CLI's: the deck being built
    // is what says which workspace is meant, so `pptx-gen build --script
    // ~/decks/projects/x/build.ts` works from anywhere.
    const workspace = workspaceFor(command, path.dirname(scriptPath));

    // Importing design.js already applied the design of the workspace found in
    // the environment. If --workspace names a different one, its design has to
    // replace that — and it has to happen HERE, before the dynamic import
    // below, because a build script's top-level `const { LM } = LAYOUT` keeps
    // whatever it sees when its module body runs.
    const patch = readDesignFileSync(workspace.designPath);
    if (patch) applyDesign(patch);

    // So that a `new Presentation()` inside the script resolves the same
    // workspace, whatever directory the script lives in.
    process.env.PPTX_GEN_WORKSPACE = workspace.root;

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

program
  .command("where")
  .description("Print the install directory")
  .action(() => {
    console.log(installDir());
  });

try {
  await program.parseAsync(process.argv);
} catch (error) {
  // A workspace or config problem is a user-facing message, not a stack trace.
  console.error((error as Error).message);
  process.exitCode = 1;
}
