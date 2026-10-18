// One segment of an area as the shapes that fill it: a rectangle from the
// baseline to the lower end plus a triangle for the rise or fall (tip at the
// higher end), a zero crossing as two triangles, the part under the baseline
// mirrored. The seam: a rectangle reaches up into its triangle and across into
// the next stretch, a triangle across on its tip side (the exactness and the
// coverage are in the tile suite). Pure geometry: slide points, y growing down.

import { describe, expect, it } from "vitest";
import {
  AREA_SEAM,
  slopePieces,
  type Piece,
  type PieceOptions,
  type Slope,
} from "./chart-shapes-area-pieces";
import type { Rect } from "./chart-shapes";
import type { Triangle } from "./chart-shapes-types";

const COLOR = "#2EC4B6";
const OPTIONS: PieceOptions = {
  id: "0.3",
  color: COLOR,
  seam: 0,
  first: true,
  last: true,
};

function slope(from: number, to: number, base = 100, width = 40): Slope {
  return { left: 0, right: width, from, to, base };
}

const rects = (out: Piece[]): Rect[] =>
  out.filter((one): one is Rect => one.kind === "rect");
const triangles = (out: Piece[]): Triangle[] =>
  out.filter((one): one is Triangle => one.kind === "triangle");

describe("slopePieces above the baseline", () => {
  it("fills a rise with a rectangle to the lower end and a triangle tipped at the right", () => {
    const out = slopePieces(slope(80, 40), OPTIONS);
    expect(rects(out).map((one) => one.box)).toEqual([
      { left: 0, top: 80, width: 40, height: 20 },
    ]);
    expect(triangles(out)).toEqual([
      {
        kind: "triangle",
        box: { left: 0, top: 40, width: 40, height: 40 },
        apex: 1,
        down: false,
        color: COLOR,
        name: "slope 0.3",
      },
    ]);
  });

  it("fills a fall with the same rectangle and a triangle tipped at the left", () => {
    const out = slopePieces(slope(40, 80), OPTIONS);
    expect(rects(out).map((one) => one.box)).toEqual([
      { left: 0, top: 80, width: 40, height: 20 },
    ]);
    expect(triangles(out).map((one) => [one.box, one.apex, one.down])).toEqual([
      [{ left: 0, top: 40, width: 40, height: 40 }, 0, false],
    ]);
  });

  it("draws a flat stretch as a rectangle alone", () => {
    const out = slopePieces(slope(70, 70), OPTIONS);
    expect(out).toHaveLength(1);
    expect(rects(out)[0]!.box).toEqual({
      left: 0,
      top: 70,
      width: 40,
      height: 30,
    });
  });

  it("draws nothing for a stretch that lies on the baseline", () => {
    expect(slopePieces(slope(100, 100), OPTIONS)).toEqual([]);
  });

  it("starts a rise from the baseline with a triangle and no rectangle", () => {
    const out = slopePieces(slope(100, 60), OPTIONS);
    expect(rects(out)).toEqual([]);
    expect(triangles(out).map((one) => [one.box, one.apex])).toEqual([
      [{ left: 0, top: 60, width: 40, height: 40 }, 1],
    ]);
  });

  it("names the rectangle area and the triangle slope, under the segment's id", () => {
    const out = slopePieces(slope(80, 40), OPTIONS);
    expect(out.map((one) => one.name)).toEqual(["area 0.3", "slope 0.3"]);
    expect(out.every((one) => one.color === COLOR)).toBe(true);
  });

  it("paints the rectangle first, so the triangle sits on top of it", () => {
    expect(slopePieces(slope(80, 40), OPTIONS).map((one) => one.kind)).toEqual([
      "rect",
      "triangle",
    ]);
  });
});

describe("slopePieces under the baseline", () => {
  it("hangs the rectangle from the baseline and the triangle under it, tip down", () => {
    const out = slopePieces(slope(120, 160), OPTIONS);
    expect(rects(out).map((one) => one.box)).toEqual([
      { left: 0, top: 100, width: 40, height: 20 },
    ]);
    expect(triangles(out).map((one) => [one.box, one.apex, one.down])).toEqual([
      [{ left: 0, top: 120, width: 40, height: 40 }, 1, true],
    ]);
  });

  it("tips the hanging triangle at the end that reaches deeper", () => {
    const out = slopePieces(slope(160, 120), OPTIONS);
    expect(triangles(out).map((one) => [one.apex, one.down])).toEqual([
      [0, true],
    ]);
  });
});

describe("slopePieces across the baseline", () => {
  it("splits a fall through zero into a triangle above and one below, meeting at the crossing", () => {
    const out = slopePieces(slope(80, 120), OPTIONS);
    expect(rects(out)).toEqual([]);
    expect(triangles(out).map((one) => [one.box, one.apex, one.down])).toEqual([
      [{ left: 0, top: 80, width: 20, height: 20 }, 0, false],
      [{ left: 20, top: 100, width: 20, height: 20 }, 1, true],
    ]);
    expect(out.map((one) => one.name)).toEqual(["slope 0.3.0", "slope 0.3.1"]);
  });

  it("splits a rise through zero the other way round", () => {
    const out = slopePieces(slope(130, 90), OPTIONS);
    expect(triangles(out).map((one) => [one.box, one.apex, one.down])).toEqual([
      [{ left: 0, top: 100, width: 30, height: 30 }, 0, true],
      [{ left: 30, top: 90, width: 10, height: 10 }, 1, false],
    ]);
  });

  it("puts the crossing where the line meets the baseline, nearer the end that is closer", () => {
    const out = triangles(slopePieces(slope(70, 110), OPTIONS));
    // 30 above at the left, 10 below at the right: it crosses 3/4 of the way.
    expect(out[0]!.box.width).toBeCloseTo(30);
    expect(out[1]!.box.left).toBeCloseTo(30);
    expect(out[1]!.box.width).toBeCloseTo(10);
  });

  it("is a triangle on one side and none on the other when the line only touches zero", () => {
    const out = slopePieces(slope(100, 130), OPTIONS);
    expect(rects(out)).toEqual([]);
    expect(triangles(out).map((one) => [one.apex, one.down])).toEqual([
      [1, true],
    ]);
  });
});

describe("slopePieces and the seam", () => {
  const seamed: PieceOptions = { ...OPTIONS, seam: AREA_SEAM };

  it("is half a point", () => {
    expect(AREA_SEAM).toBe(0.5);
  });

  it("lets the rectangle reach up into the triangle and across into the next segment", () => {
    const out = slopePieces(slope(80, 40), { ...seamed, last: false });
    expect(rects(out)[0]!.box).toEqual({
      left: 0,
      top: 79.5,
      width: 40.5,
      height: 20.5,
    });
  });

  it("does not reach past the last segment of a run", () => {
    const out = slopePieces(slope(80, 40), seamed);
    expect(rects(out)[0]!.box.width).toBe(40);
    expect(rects(out)[0]!.box.top).toBe(79.5);
  });

  it("never lets the rectangle climb higher than the triangle's tip", () => {
    // A rise of 0.2 points: the rectangle may overlap by 0.2, not by 0.5.
    const out = slopePieces(slope(80, 79.8), seamed);
    expect(rects(out)[0]!.box.top).toBeCloseTo(79.8);
  });

  it("leaves out a triangle no taller than the seam: the rectangle reached up to its tip", () => {
    for (const rise of [0.2, AREA_SEAM]) {
      const out = slopePieces(slope(80, 80 - rise), seamed);
      expect(triangles(out)).toEqual([]);
      expect(rects(out)[0]!.box.top).toBeCloseTo(80 - rise);
    }
    // Under the baseline as well.
    const hanging = slopePieces(slope(120, 120.4), seamed);
    expect(triangles(hanging)).toEqual([]);
    expect(rects(hanging)[0]!.box.height).toBeCloseTo(20.4);
  });

  it("keeps a triangle taller than the seam, and one with no rectangle to cover it", () => {
    expect(triangles(slopePieces(slope(80, 79.4), seamed))).toHaveLength(1);
    // A rise from the baseline has no rectangle: the sliver is all there is.
    const sliver = slopePieces(slope(100, 99.7), seamed);
    expect(rects(sliver)).toEqual([]);
    expect(triangles(sliver)).toHaveLength(1);
  });

  it("reaches down into the triangle under the baseline", () => {
    const out = slopePieces(slope(120, 160), { ...seamed, last: false });
    expect(rects(out)[0]!.box).toEqual({
      left: 0,
      top: 100,
      width: 40.5,
      height: 20.5,
    });
  });

  it("leaves a triangle exactly where the line is when nothing lies on its tip's side", () => {
    const out = slopePieces(slope(80, 40), seamed);
    expect(triangles(out)[0]).toMatchObject({
      box: { left: 0, top: 40, width: 40, height: 40 },
      apex: 1,
    });
    const fall = slopePieces(slope(40, 80), seamed);
    expect(triangles(fall)[0]).toMatchObject({
      box: { left: 0, top: 40, width: 40, height: 40 },
      apex: 0,
    });
  });

  it("widens a rising triangle to the right when a stretch follows, the tip staying on the line", () => {
    const out = slopePieces(slope(80, 40), { ...seamed, last: false });
    const [rise] = triangles(out);
    expect(rise!.box).toEqual({ left: 0, top: 40, width: 40.5, height: 40 });
    expect(rise!.apex).toBeCloseTo(40 / 40.5);
    expect(rise!.box.left + rise!.apex * rise!.box.width).toBeCloseTo(40);
    // A stretch before it is no reason to reach left: its tip is on the right.
    expect(
      triangles(slopePieces(slope(80, 40), { ...seamed, first: false }))[0]!.box
        .left,
    ).toBe(0);
  });

  it("widens a falling triangle to the left when a stretch precedes it", () => {
    const out = slopePieces(slope(40, 80), { ...seamed, first: false });
    const [fall] = triangles(out);
    expect(fall!.box).toEqual({ left: -0.5, top: 40, width: 40.5, height: 40 });
    expect(fall!.apex).toBeCloseTo(0.5 / 40.5);
    expect(fall!.box.left + fall!.apex * fall!.box.width).toBeCloseTo(0);
    expect(
      triangles(slopePieces(slope(40, 80), { ...seamed, last: false }))[0]!.box
        .width,
    ).toBe(40);
  });

  it("widens the hanging triangle and the two across the baseline the same way, on their tip sides", () => {
    const hanging = slopePieces(slope(120, 160), { ...seamed, last: false });
    expect(triangles(hanging)[0]).toMatchObject({
      box: { left: 0, width: 40.5 },
      down: true,
    });
    const crossed = slopePieces(slope(80, 120), {
      ...seamed,
      first: false,
      last: false,
    });
    const [above, below] = triangles(crossed);
    expect(above!.box.left).toBe(-0.5);
    expect(above!.box.left + above!.apex * above!.box.width).toBeCloseTo(0);
    expect(below!.box.left).toBe(20);
    expect(below!.box.width).toBeCloseTo(20.5);
    expect(below!.box.left + below!.apex * below!.box.width).toBeCloseTo(40);
  });
});
