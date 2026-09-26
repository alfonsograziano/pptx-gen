import path from "node:path";
import { fileURLToPath } from "node:url";
import { writeFile } from "node:fs/promises";
import { deflateSync } from "node:zlib";
import type { ShotFn, ShotRequest } from "./html-shot.js";

const HERE = path.dirname(fileURLToPath(import.meta.url));

/**
 * The install's seed template library, which the tests use as their fixture
 * library. Kept here rather than re-derived in each test file so moving it is a
 * one-line change.
 */
export const STARTER_TEMPLATES = path.resolve(HERE, "..", "starter", "templates");

/** The install's bundled assets, including the icon set. */
export const BUNDLED_ASSETS = path.resolve(HERE, "..", "assets");

/**
 * A 1x1 transparent PNG. Lets tests exercise image and figure handling without
 * depending on a binary asset in the repo or on a browser being installed.
 */
export const TINY_PNG = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC",
  "base64"
);

export type FakeShot = ShotFn & {
  /** How many times the rasterizer actually ran, for asserting on caching. */
  calls: number;
  requests: ShotRequest[];
};

/**
 * A stand-in rasterizer that writes a real PNG without launching a browser.
 *
 * `pxWidth`/`pxHeight` are reported as the requested viewport times the scale,
 * matching what Chrome does, so fitting maths can be tested honestly.
 */
export function fakeShot(options: { fail?: boolean } = {}): FakeShot {
  const shot = (async (request: ShotRequest) => {
    shot.calls += 1;
    shot.requests.push(request);
    if (options.fail) throw new Error("fake rasterizer failure");
    await writeFile(request.pngPath, TINY_PNG);
    return { pxWidth: request.width * request.scale, pxHeight: request.height * request.scale };
  }) as FakeShot;
  shot.calls = 0;
  shot.requests = [];
  return shot;
}

/**
 * A valid grey PNG of any size: real IHDR, CRCs and deflated pixel rows. For
 * tests whose maths depends on a picture's proportions, and that LibreOffice
 * should still be able to render.
 */
export function makePng(width: number, height: number, shade = 0x88): Buffer {
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header[8] = 8; // bit depth
  header[9] = 0; // greyscale
  const row = Buffer.alloc(width + 1, shade);
  row[0] = 0; // no filter
  const pixels = deflateSync(Buffer.concat(Array.from({ length: height }, () => row)));
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", header),
    pngChunk("IDAT", pixels),
    pngChunk("IEND", Buffer.alloc(0))
  ]);
}

function pngChunk(type: string, data: Buffer): Buffer {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([length, body, crc]);
}

function crc32(bytes: Buffer): number {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
