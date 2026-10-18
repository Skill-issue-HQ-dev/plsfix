// The pieces of one area stretch against the trapezoid under the line, as
// drawn: with no seam they add up to it exactly, with the seam they cover every
// point of it and reach past it by the seam along each axis at most (a corner:
// the seam times the square root of two), whatever the slope, the side of the
// baseline and the neighbours of the stretch (fast-check, fixed seed). Pure
// geometry: slide points, y growing downwards.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  AREA_SEAM,
  slopePieces,
  type Piece,
  type PieceOptions,
  type Slope,
} from "./chart-shapes-area-pieces";
import type { Triangle } from "./chart-shapes-types";

const OPTIONS: PieceOptions = {
  id: "0.3",
  color: "#2EC4B6",
  seam: 0,
  first: true,
  last: true,
};

function slope(from: number, to: number, base = 100, width = 40): Slope {
  return { left: 0, right: width, from, to, base };
}

type Point = readonly [number, number];

// The trapezoid between the baseline and the line, as the convex pieces it is
// made of (one, or two triangles where the line crosses the baseline).
function polygons(one: Slope): Point[][] {
  const { left, right, from, to, base } = one;
  const crosses = (from - base) * (to - base) < 0;
  if (!crosses) {
    return [
      [
        [left, base],
        [right, base],
        [right, to],
        [left, from],
      ],
    ];
  }
  const cross = left + ((right - left) * (base - from)) / (to - from);
  return [
    [
      [left, base],
      [cross, base],
      [left, from],
    ],
    [
      [cross, base],
      [right, base],
      [right, to],
    ],
  ];
}

function trianglePoints(one: Triangle): Point[] {
  const { left, top, width, height } = one.box;
  const tipX = left + one.apex * width;
  return one.down
    ? [
        [left, top],
        [left + width, top],
        [tipX, top + height],
      ]
    : [
        [left, top + height],
        [left + width, top + height],
        [tipX, top],
      ]; // standing on its bottom edge
}

function inConvex(points: readonly Point[], x: number, y: number): boolean {
  let sign = 0;
  for (let i = 0; i < points.length; i += 1) {
    const [ax, ay] = points[i]!;
    const [bx, by] = points[(i + 1) % points.length]!;
    const cross = (bx - ax) * (y - ay) - (by - ay) * (x - ax);
    if (Math.abs(cross) < 1e-9) continue;
    if (sign === 0) sign = Math.sign(cross);
    else if (sign !== Math.sign(cross)) return false;
  }
  return true;
}

function distance(points: readonly Point[], x: number, y: number): number {
  if (inConvex(points, x, y)) return 0;
  let best = Infinity;
  for (let i = 0; i < points.length; i += 1) {
    const [ax, ay] = points[i]!;
    const [bx, by] = points[(i + 1) % points.length]!;
    const length2 = (bx - ax) ** 2 + (by - ay) ** 2;
    const t =
      length2 === 0
        ? 0
        : Math.max(
            0,
            Math.min(
              1,
              ((x - ax) * (bx - ax) + (y - ay) * (by - ay)) / length2,
            ),
          );
    best = Math.min(
      best,
      Math.hypot(x - ax - t * (bx - ax), y - ay - t * (by - ay)),
    );
  }
  return best;
}

function covers(out: Piece[], x: number, y: number): boolean {
  return out.some((one) => {
    if (one.kind === "triangle") return inConvex(trianglePoints(one), x, y);
    const { left, top, width, height } = one.box;
    return x >= left && x <= left + width && y >= top && y <= top + height;
  });
}

function areaOf(out: Piece[]): number {
  return out.reduce((sum, one) => {
    const { width, height } = one.box;
    return (
      sum + (one.kind === "triangle" ? (width * height) / 2 : width * height)
    );
  }, 0);
}

// How far a piece reaches past the area: the seam along each axis at once,
// which is the seam times the square root of two at the corner of a rectangle
// that reached up into its triangle and across into the next stretch.
const CORNER_REACH = AREA_SEAM * Math.SQRT2;

function vertices(piece: Piece): Point[] {
  if (piece.kind === "triangle") return trianglePoints(piece);
  const { left, top, width, height } = piece.box;
  return [
    [left, top],
    [left + width, top],
    [left + width, top + height],
    [left, top + height],
  ];
}

// The distance to a convex area is convex, so a convex piece reaches farthest
// at one of its vertices: no sampling grid can miss it. A crossing stretch has
// two parts, each piece measured against its own (`.0`, `.1`).
function farthest(out: Piece[], parts: Point[][]): number {
  return Math.max(
    ...out.flatMap((piece) => {
      const part = parts[piece.name.endsWith(".1") ? 1 : 0]!;
      return vertices(piece).map(([x, y]) => distance(part, x, y));
    }),
  );
}

// Tenths of a point, so a rise under the seam (the stretches that have no
// triangle of their own) comes up as often as any other.
const slopes = fc
  .tuple(
    fc.integer({ min: 0, max: 2000 }),
    fc.integer({ min: 0, max: 2000 }),
    fc.integer({ min: 400, max: 1600 }),
    fc.integer({ min: 40, max: 600 }),
  )
  .map(([from, to, base, width]) =>
    slope(from / 10, to / 10, base / 10, width / 10),
  );

describe("slopePieces tile the area under the line", () => {
  it("add up to exactly the trapezoid when there is no seam", () => {
    fc.assert(
      fc.property(slopes, (one) => {
        const exact = polygons(one).reduce((sum, polygon) => {
          const [a, b, c, d] = polygon;
          // Shoelace over the pieces the trapezoid is made of.
          const points = d === undefined ? [a!, b!, c!] : [a!, b!, c!, d];
          let twice = 0;
          points.forEach(([x, y], i) => {
            const [nx, ny] = points[(i + 1) % points.length]!;
            twice += x * ny - nx * y;
          });
          return sum + Math.abs(twice) / 2;
        }, 0);
        expect(areaOf(slopePieces(one, OPTIONS))).toBeCloseTo(exact, 6);
      }),
      { seed: 20261003, numRuns: 500 },
    );
  });

  it("cover every point of the area, whatever the slope", () => {
    fc.assert(
      fc.property(slopes, fc.boolean(), fc.boolean(), (one, first, last) => {
        const out = slopePieces(one, {
          ...OPTIONS,
          seam: AREA_SEAM,
          first,
          last,
        });
        const parts = polygons(one);
        const { left, right, base } = one;
        const lowest = Math.min(one.from, one.to, base) - 1;
        const highest = Math.max(one.from, one.to, base) + 1;
        for (let x = left - 1; x <= right + 1; x += 0.7) {
          for (let y = lowest; y <= highest; y += 0.9) {
            // Every point of the area is covered, margin for the edges.
            const deep = parts.some(
              (part) =>
                inConvex(part, x, y) &&
                [
                  [0.05, 0],
                  [-0.05, 0],
                  [0, 0.05],
                  [0, -0.05],
                ].every(([dx, dy]) => inConvex(part, x + dx!, y + dy!)),
            );
            if (deep) {
              expect(covers(out, x, y), `${String(x)},${String(y)}`).toBe(true);
            }
          }
        }
      }),
      { seed: 20261003, numRuns: 60 },
    );
  });

  it("reach past the area by the seam along each axis at most, whatever the slope", () => {
    fc.assert(
      fc.property(slopes, fc.boolean(), fc.boolean(), (one, first, last) => {
        const out = slopePieces(one, {
          ...OPTIONS,
          seam: AREA_SEAM,
          first,
          last,
        });
        expect(farthest(out, polygons(one))).toBeLessThanOrEqual(
          CORNER_REACH + 1e-9,
        );
      }),
      { seed: 20261003, numRuns: 500 },
    );
  });

  it("are the rectangle alone, holding the whole trapezoid, when the rise is within the seam", () => {
    const within = fc.tuple(
      fc.integer({ min: 10, max: 900 }),
      fc.integer({ min: 0, max: 5 }),
      fc.boolean(),
      fc.boolean(),
      fc.boolean(),
      fc.boolean(),
    );
    fc.assert(
      fc.property(within, ([near, rise, falling, hanging, first, last]) => {
        // Tenths of a point: the lower end `near` away from the baseline, the
        // higher one `rise` farther, above it or (hanging) below.
        const side = hanging ? 1 : -1;
        const one = slope(
          100 + (side * near) / 10,
          100 + (side * (near + (falling ? rise : -rise))) / 10,
        );
        const out = slopePieces(one, {
          ...OPTIONS,
          seam: AREA_SEAM,
          first,
          last,
        });
        expect(out).toHaveLength(1);
        const [box] = out.map((piece) => piece.box);
        for (const [x, y] of polygons(one).flat()) {
          expect(x).toBeGreaterThanOrEqual(box!.left - 1e-9);
          expect(x).toBeLessThanOrEqual(box!.left + box!.width + 1e-9);
          expect(y).toBeGreaterThanOrEqual(box!.top - 1e-9);
          expect(y).toBeLessThanOrEqual(box!.top + box!.height + 1e-9);
        }
      }),
      { seed: 20261003, numRuns: 300 },
    );
  });

  it("reach the seam times the square root of two at the corner of a rectangle with a stretch after it", () => {
    // A fall: its rectangle reaches up half a point past the line's end and
    // half a point across into the next stretch.
    const one = slope(60, 80);
    const out = slopePieces(one, { ...OPTIONS, seam: AREA_SEAM, last: false });
    expect(farthest(out, polygons(one))).toBeCloseTo(CORNER_REACH, 9);
  });

  it("never makes a side negative, or a piece smaller than a point's hundredth", () => {
    fc.assert(
      fc.property(slopes, fc.boolean(), (one, last) => {
        for (const piece of slopePieces(one, {
          ...OPTIONS,
          seam: AREA_SEAM,
          last,
        })) {
          expect(piece.box.width).toBeGreaterThanOrEqual(0.01);
          expect(piece.box.height).toBeGreaterThanOrEqual(0.01);
        }
      }),
      { seed: 20261003, numRuns: 500 },
    );
  });
});
