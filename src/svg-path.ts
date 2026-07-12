// Convert SVG vector art (paths, circles, simple primitives) into flattened
// polyline/curve subpaths, so it can be emitted as native PowerPoint custom
// geometry (editable + recolorable in PowerPoint AND Google Slides) instead of
// a rasterized image. Curves are flattened to line segments, which keeps the
// converter small and robust while staying visually smooth at icon scale.

export type Pt = [number, number];
export type SubPath = { pts: Pt[]; closed: boolean };
export type ParsedSvg = { vbW: number; vbH: number; subs: SubPath[] };

function sampleCubic(out: Pt[], x0: number, y0: number, x1: number, y1: number, x2: number, y2: number, x3: number, y3: number, n = 12): void {
  for (let k = 1; k <= n; k += 1) {
    const t = k / n;
    const mt = 1 - t;
    const a = mt * mt * mt;
    const b = 3 * mt * mt * t;
    const c = 3 * mt * t * t;
    const d = t * t * t;
    out.push([a * x0 + b * x1 + c * x2 + d * x3, a * y0 + b * y1 + c * y2 + d * y3]);
  }
}

function sampleQuad(out: Pt[], x0: number, y0: number, x1: number, y1: number, x2: number, y2: number, n = 10): void {
  for (let k = 1; k <= n; k += 1) {
    const t = k / n;
    const mt = 1 - t;
    out.push([mt * mt * x0 + 2 * mt * t * x1 + t * t * x2, mt * mt * y0 + 2 * mt * t * y1 + t * t * y2]);
  }
}

function sampleArc(out: Pt[], x0: number, y0: number, rxIn: number, ryIn: number, phiDeg: number, laf: number, sf: number, x: number, y: number, n = 14): void {
  let rx = Math.abs(rxIn);
  let ry = Math.abs(ryIn);
  if (rx === 0 || ry === 0) {
    out.push([x, y]);
    return;
  }
  const phi = (phiDeg * Math.PI) / 180;
  const cp = Math.cos(phi);
  const sp = Math.sin(phi);
  const dx = (x0 - x) / 2;
  const dy = (y0 - y) / 2;
  const x1p = cp * dx + sp * dy;
  const y1p = -sp * dx + cp * dy;

  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    const s = Math.sqrt(lambda);
    rx *= s;
    ry *= s;
  }

  const sign = laf !== sf ? 1 : -1;
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  const co = sign * Math.sqrt(Math.max(0, num / den));
  const cxp = (co * rx * y1p) / ry;
  const cyp = (-co * ry * x1p) / rx;
  const cx = cp * cxp - sp * cyp + (x0 + x) / 2;
  const cy = sp * cxp + cp * cyp + (y0 + y) / 2;

  const angle = (ux: number, uy: number, vx: number, vy: number): number => {
    const dot = ux * vx + uy * vy;
    const len = Math.hypot(ux, uy) * Math.hypot(vx, vy);
    let a = Math.acos(Math.min(1, Math.max(-1, dot / len)));
    if (ux * vy - uy * vx < 0) a = -a;
    return a;
  };

  const th1 = angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dth = angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sf && dth > 0) dth -= 2 * Math.PI;
  if (sf && dth < 0) dth += 2 * Math.PI;

  for (let k = 1; k <= n; k += 1) {
    const t = th1 + dth * (k / n);
    const ex = cx + rx * Math.cos(t) * cp - ry * Math.sin(t) * sp;
    const ey = cy + rx * Math.cos(t) * sp + ry * Math.sin(t) * cp;
    out.push([ex, ey]);
  }
}

/** Flatten a single SVG path "d" string into absolute subpaths. */
export function flattenPath(d: string): SubPath[] {
  const subs: SubPath[] = [];
  let cur: Pt[] = [];
  let closed = false;
  let x = 0;
  let y = 0;
  let startX = 0;
  let startY = 0;
  let prevCtrl: Pt | null = null;
  let prevCmd = "";

  const tokens = d.match(/[a-zA-Z]|-?\d*\.?\d+(?:e[-+]?\d+)?/g) ?? [];
  let i = 0;
  const num = (): number => Number.parseFloat(tokens[i++]);
  const isCmd = (t: string): boolean => /[a-zA-Z]/.test(t);

  const pushCur = (): void => {
    if (cur.length) subs.push({ pts: cur, closed });
    cur = [];
    closed = false;
  };

  let cmd = "";
  while (i < tokens.length) {
    if (isCmd(tokens[i])) cmd = tokens[i++];
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();

    if (C === "M") {
      pushCur();
      let nx = num();
      let ny = num();
      if (rel) {
        nx += x;
        ny += y;
      }
      x = nx;
      y = ny;
      startX = x;
      startY = y;
      cur.push([x, y]);
      cmd = rel ? "l" : "L"; // implicit lineTo for subsequent pairs
    } else if (C === "L") {
      let nx = num();
      let ny = num();
      if (rel) {
        nx += x;
        ny += y;
      }
      x = nx;
      y = ny;
      cur.push([x, y]);
    } else if (C === "H") {
      let nx = num();
      if (rel) nx += x;
      x = nx;
      cur.push([x, y]);
    } else if (C === "V") {
      let ny = num();
      if (rel) ny += y;
      y = ny;
      cur.push([x, y]);
    } else if (C === "C") {
      let x1 = num();
      let y1 = num();
      let x2 = num();
      let y2 = num();
      let nx = num();
      let ny = num();
      if (rel) {
        x1 += x; y1 += y; x2 += x; y2 += y; nx += x; ny += y;
      }
      sampleCubic(cur, x, y, x1, y1, x2, y2, nx, ny);
      prevCtrl = [x2, y2];
      x = nx;
      y = ny;
    } else if (C === "S") {
      let x2 = num();
      let y2 = num();
      let nx = num();
      let ny = num();
      if (rel) {
        x2 += x; y2 += y; nx += x; ny += y;
      }
      const reflect: boolean = (prevCmd === "C" || prevCmd === "S") && prevCtrl !== null;
      const x1: number = reflect ? 2 * x - (prevCtrl as Pt)[0] : x;
      const y1: number = reflect ? 2 * y - (prevCtrl as Pt)[1] : y;
      sampleCubic(cur, x, y, x1, y1, x2, y2, nx, ny);
      prevCtrl = [x2, y2];
      x = nx;
      y = ny;
    } else if (C === "Q") {
      let x1 = num();
      let y1 = num();
      let nx = num();
      let ny = num();
      if (rel) {
        x1 += x; y1 += y; nx += x; ny += y;
      }
      sampleQuad(cur, x, y, x1, y1, nx, ny);
      prevCtrl = [x1, y1];
      x = nx;
      y = ny;
    } else if (C === "T") {
      let nx = num();
      let ny = num();
      if (rel) {
        nx += x; ny += y;
      }
      const reflect: boolean = (prevCmd === "Q" || prevCmd === "T") && prevCtrl !== null;
      const x1: number = reflect ? 2 * x - (prevCtrl as Pt)[0] : x;
      const y1: number = reflect ? 2 * y - (prevCtrl as Pt)[1] : y;
      sampleQuad(cur, x, y, x1, y1, nx, ny);
      prevCtrl = [x1, y1];
      x = nx;
      y = ny;
    } else if (C === "A") {
      const rx = num();
      const ry = num();
      const rot = num();
      const laf = num();
      const sf = num();
      let nx = num();
      let ny = num();
      if (rel) {
        nx += x; ny += y;
      }
      sampleArc(cur, x, y, rx, ry, rot, laf, sf, nx, ny);
      x = nx;
      y = ny;
    } else if (C === "Z") {
      cur.push([startX, startY]);
      closed = true;
      pushCur();
      x = startX;
      y = startY;
    } else {
      i += 1; // skip anything unrecognised
    }

    prevCmd = C;
  }

  pushCur();
  return subs;
}

function circleSub(cx: number, cy: number, r: number, n = 28): SubPath {
  const pts: Pt[] = [];
  for (let k = 0; k <= n; k += 1) {
    const t = (2 * Math.PI * k) / n;
    pts.push([cx + r * Math.cos(t), cy + r * Math.sin(t)]);
  }
  return { pts, closed: true };
}

function attr(tag: string, name: string): number | undefined {
  const m = tag.match(new RegExp(`\\b${name}="([\\d.eE+-]+)"`));
  return m ? Number.parseFloat(m[1]) : undefined;
}

/** Parse an SVG string into a viewBox and a flat list of subpaths. */
export function parseSvg(svg: string): ParsedSvg {
  let vbW = 24;
  let vbH = 24;
  const vb = svg.match(/viewBox="([\d.\s-]+)"/);
  if (vb) {
    const parts = vb[1].trim().split(/\s+/).map(Number);
    if (parts.length === 4) {
      vbW = parts[2];
      vbH = parts[3];
    }
  } else {
    const w = svg.match(/\bwidth="([\d.]+)"/);
    const h = svg.match(/\bheight="([\d.]+)"/);
    if (w) vbW = Number.parseFloat(w[1]);
    if (h) vbH = Number.parseFloat(h[1]);
  }

  const subs: SubPath[] = [];

  for (const m of svg.matchAll(/<path\b[^>]*\bd="([^"]+)"[^>]*>/g)) {
    subs.push(...flattenPath(m[1]));
  }
  for (const m of svg.matchAll(/<circle\b[^>]*>/g)) {
    const cx = attr(m[0], "cx");
    const cy = attr(m[0], "cy");
    const r = attr(m[0], "r");
    if (cx !== undefined && cy !== undefined && r !== undefined) subs.push(circleSub(cx, cy, r));
  }
  for (const m of svg.matchAll(/<line\b[^>]*>/g)) {
    const x1 = attr(m[0], "x1");
    const y1 = attr(m[0], "y1");
    const x2 = attr(m[0], "x2");
    const y2 = attr(m[0], "y2");
    if (x1 !== undefined && y1 !== undefined && x2 !== undefined && y2 !== undefined) {
      subs.push({ pts: [[x1, y1], [x2, y2]], closed: false });
    }
  }
  for (const m of svg.matchAll(/<rect\b[^>]*>/g)) {
    const rx = attr(m[0], "x") ?? 0;
    const ry = attr(m[0], "y") ?? 0;
    const rw = attr(m[0], "width");
    const rh = attr(m[0], "height");
    if (rw !== undefined && rh !== undefined) {
      subs.push({ pts: [[rx, ry], [rx + rw, ry], [rx + rw, ry + rh], [rx, ry + rh], [rx, ry]], closed: true });
    }
  }
  for (const m of svg.matchAll(/<(?:polyline|polygon)\b[^>]*\bpoints="([^"]+)"[^>]*>/g)) {
    const nums = (m[1].match(/-?\d*\.?\d+/g) ?? []).map(Number);
    const pts: Pt[] = [];
    for (let k = 0; k + 1 < nums.length; k += 2) pts.push([nums[k], nums[k + 1]]);
    if (pts.length) subs.push({ pts, closed: /<polygon/.test(m[0]) });
  }

  return { vbW, vbH, subs };
}

type GeomPoint =
  | { x: number; y: number; moveTo?: boolean }
  | { close: true };

/**
 * Convert parsed SVG subpaths to PptxGenJS custom-geometry points, scaled into a
 * box of width/height `w`/`h` (inches). Coordinates are relative to the shape
 * origin (PptxGenJS renders custGeom points at shape offset + point).
 */
export function svgToGeomPoints(parsed: ParsedSvg, w: number, h: number): GeomPoint[] {
  const sx = w / parsed.vbW;
  const sy = h / parsed.vbH;
  const r = (n: number): number => Math.round(n * 10000) / 10000;
  const points: GeomPoint[] = [];
  for (const sub of parsed.subs) {
    sub.pts.forEach((p, idx) => {
      const gx = r(p[0] * sx);
      const gy = r(p[1] * sy);
      points.push(idx === 0 ? { x: gx, y: gy, moveTo: true } : { x: gx, y: gy });
    });
    if (sub.closed) points.push({ close: true });
  }
  return points;
}
