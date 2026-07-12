import { lstat, mkdir, readdir, readlink, rm, symlink } from "node:fs/promises";
import path from "node:path";

const templateRoot = path.resolve("templates");
const previewRoot = path.resolve("templates-preview");

type Result = {
  created: number;
  removed: number;
  skipped: string[];
  missingScreenshots: string[];
};

await mkdir(previewRoot, { recursive: true });

const result: Result = {
  created: 0,
  removed: 0,
  skipped: [],
  missingScreenshots: []
};

const templates = (await readdir(templateRoot, { withFileTypes: true }))
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();

const expectedLinks = new Set(templates.map((templateName) => `${templateName}.png`));
for (const entry of await readdir(previewRoot, { withFileTypes: true })) {
  const previewPath = path.join(previewRoot, entry.name);
  const stat = await lstat(previewPath);
  if (stat.isSymbolicLink() && !expectedLinks.has(entry.name)) {
    await rm(previewPath);
    result.removed += 1;
  }
}

for (const templateName of templates) {
  const screenshot = path.join(templateRoot, templateName, "screenshots", "slide-01.png");
  if (!await exists(screenshot)) {
    result.missingScreenshots.push(templateName);
    continue;
  }

  const linkPath = path.join(previewRoot, `${templateName}.png`);
  const relativeTarget = path.relative(previewRoot, screenshot);
  const existing = await lstat(linkPath).catch(() => undefined);

  if (existing?.isSymbolicLink()) {
    const currentTarget = await readlink(linkPath);
    if (currentTarget !== relativeTarget) {
      await rm(linkPath);
      await symlink(relativeTarget, linkPath);
    }
    result.created += 1;
    continue;
  }

  if (existing) {
    result.skipped.push(`${templateName}.png`);
    continue;
  }

  await symlink(relativeTarget, linkPath);
  result.created += 1;
}

console.log(`Synced ${result.created} template preview link(s) in ${path.relative(process.cwd(), previewRoot)}.`);

if (result.removed > 0) {
  console.log(`Removed ${result.removed} stale preview link(s).`);
}

if (result.missingScreenshots.length > 0) {
  console.log(`Missing screenshots: ${result.missingScreenshots.join(", ")}`);
}

if (result.skipped.length > 0) {
  console.log(`Skipped existing non-symlink file(s): ${result.skipped.join(", ")}`);
}

async function exists(filePath: string): Promise<boolean> {
  return lstat(filePath).then(() => true, () => false);
}
