import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

export class PptxPackage {
  private constructor(private readonly zip: JSZip) {}

  static async load(filePath: string): Promise<PptxPackage> {
    const buffer = await readFile(filePath);
    return new PptxPackage(await JSZip.loadAsync(buffer));
  }

  async save(filePath: string): Promise<void> {
    const buffer = await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE"
    });
    await writeFile(filePath, buffer);
  }

  has(filePath: string): boolean {
    return this.zip.file(normalizeZipPath(filePath)) !== null;
  }

  async text(filePath: string): Promise<string> {
    const file = this.zip.file(normalizeZipPath(filePath));
    if (!file) throw new Error(`PPTX entry not found: ${filePath}`);
    return file.async("text");
  }

  async bytes(filePath: string): Promise<Buffer> {
    const file = this.zip.file(normalizeZipPath(filePath));
    if (!file) throw new Error(`PPTX entry not found: ${filePath}`);
    return file.async("nodebuffer");
  }

  setText(filePath: string, value: string): void {
    this.zip.file(normalizeZipPath(filePath), value);
  }

  setBytes(filePath: string, value: Buffer | string): void {
    this.zip.file(normalizeZipPath(filePath), value);
  }

  async copy(source: string, target: string): Promise<void> {
    if (!this.has(source)) return;
    this.setBytes(target, await this.bytes(source));
  }

  remove(filePath: string): void {
    this.zip.remove(normalizeZipPath(filePath));
  }

  files(prefix = ""): string[] {
    const normalizedPrefix = normalizeZipPath(prefix);
    return Object.keys(this.zip.files).filter((file) => file.startsWith(normalizedPrefix));
  }
}

function normalizeZipPath(filePath: string): string {
  return filePath.split(path.sep).join("/");
}
