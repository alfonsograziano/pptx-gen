import { readdir } from "node:fs/promises";
import path from "node:path";
import type { FieldsFile, TemplateMetadata } from "./types.js";
import { readYamlFile } from "./fs.js";

export type TemplateRecord = {
  id: string;
  dir: string;
  pptxPath: string;
  metadataPath: string;
  metadata: TemplateMetadata;
  fieldsPath: string;
  fieldsFile: FieldsFile;
};

export async function listTemplateIds(templateRoot: string): Promise<string[]> {
  const entries = await readdir(templateRoot, { withFileTypes: true }).catch(() => []);
  return entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
}

export async function loadTemplate(templateRoot: string, templateName: string): Promise<TemplateRecord> {
  const dir = path.resolve(templateRoot, templateName);
  const pptxPath = path.join(dir, "template.pptx");
  const metadataPath = path.join(dir, "template.yml");
  const fieldsPath = path.join(dir, "fields.yml");
  const metadata = await readYamlFile<TemplateMetadata>(metadataPath);
  const fieldsFile = await readYamlFile<FieldsFile>(fieldsPath);

  return {
    id: templateName,
    dir,
    pptxPath,
    metadataPath,
    metadata,
    fieldsPath,
    fieldsFile
  };
}
