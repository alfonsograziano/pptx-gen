import { test } from "node:test";
import assert from "node:assert/strict";
import { dotToRing, flattenPath, parseSvg, svgToGeomPoints } from "./svg-path.js";

test("flattenPath turns a move+line into an absolute subpath", () => {
  const subs = flattenPath("M0 0 L10 0");
  assert.equal(subs.length, 1);
  assert.deepEqual(subs[0].pts, [
    [0, 0],
    [10, 0]
  ]);
  assert.equal(subs[0].closed, false);
});

test("flattenPath closes a Z subpath back to the start", () => {
  const subs = flattenPath("M0 0 L10 0 L10 10 Z");
  assert.equal(subs[0].closed, true);
  assert.deepEqual(subs[0].pts.at(-1), [0, 0]);
});

test("parseSvg reads the viewBox and collects primitives", () => {
  const svg = `<svg viewBox="0 0 24 24"><line x1="2" y1="2" x2="22" y2="2"/><rect x="0" y="0" width="10" height="10"/></svg>`;
  const parsed = parseSvg(svg);
  assert.equal(parsed.vbW, 24);
  assert.equal(parsed.vbH, 24);
  assert.equal(parsed.subs.length, 2);
});

test("svgToGeomPoints scales into the target box and marks the first point moveTo", () => {
  const parsed = parseSvg(`<svg viewBox="0 0 24 24"><line x1="0" y1="0" x2="24" y2="24"/></svg>`);
  const points = svgToGeomPoints(parsed, 1, 1);
  assert.equal(points.length, 2);
  assert.deepEqual(points[0], { x: 0, y: 0, moveTo: true });
  assert.deepEqual(points[1], { x: 1, y: 1 });
});

test("svgToGeomPoints turns a Lucide dot into a closed ring so it always renders", () => {
  const parsed = parseSvg(`<svg viewBox="0 0 24 24"><path d="M12 9v4"/><path d="M12 17h.01"/></svg>`);
  const points = svgToGeomPoints(parsed, 24, 24);
  // The line keeps its two points; the dot becomes a 9-point ring plus a close.
  assert.equal(points.length, 2 + 9 + 1);
  assert.deepEqual(points.at(-1), { close: true });
  const ring = points.slice(2, -1) as Array<{ x: number; y: number }>;
  for (const p of ring) assert.ok(Math.abs(Math.hypot(p.x - 12.005, p.y - 17) - 0.1) < 0.001);
});

test("dotToRing leaves real strokes alone", () => {
  const sub = {
    pts: [
      [12, 9],
      [12, 13]
    ] as Array<[number, number]>,
    closed: false
  };
  assert.equal(dotToRing(sub), sub);
});
