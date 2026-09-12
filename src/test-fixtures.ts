import path from "node:path";
import { fileURLToPath } from "node:url";
import { writeFile } from "node:fs/promises";
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
