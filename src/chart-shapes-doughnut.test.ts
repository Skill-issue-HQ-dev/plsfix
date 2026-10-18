// The doughnut layout, slice by slice: one BlockArc per slice on a square box,
// the slices cut clockwise from 12 o'clock so their sweeps close the ring
// exactly, a lone 100 % slice as one Donut, a zero or blank slice drawing
// nothing and keeping the colours' order, and the ring's thickness from the hole
// (a fraction of the doughnut's outer diameter, as Excel's doughnutHoleSize
// is). The rings, the labels and the legend are in
// chart-shapes-doughnut.rings.test.ts. Pure geometry in slide points.

import { describe, expect, it } from "vitest";
import { layoutDoughnut } from "./chart-shapes-doughnut";
import { normalizeAngle } from "./link/chart-model";
import {
  BOX,
  COLORS,
  DARK,
  arcs,
  circle,
  donuts,
  doughnut,
  labels,
  radii,
  sweep,
} from "../test/chart-doughnut.support";

describe("layoutDoughnut: the slices", () => {
  const data = doughnut([[1, 1, 2, 0]]);

  it("cuts BlockArcs clockwise from 12 o'clock, skips a zero slice and closes the ring", () => {
    const out = layoutDoughnut(data, BOX);
    expect(arcs(out).map((one) => [one.start, one.end])).toEqual([
      [-90, 0],
      [0, 90],
      [90, -90],
    ]);
    expect(arcs(out).map((one) => one.name)).toEqual([
      "slice 0.0",
      "slice 0.1",
      "slice 0.2",
    ]);
  });

  it("closes the ring exactly: the sweeps add up to 360 and each slice starts where the last ended", () => {
    for (const values of [
      [3, 7, 11, 13, 17, 19],
      [1, 999, 1],
      [0.1, 0.2, 0.3, 0.4],
      [1_000_000, 1, 1],
      [5, 4, 3, 2, 1, 0.5, 0.25, 0.125],
    ]) {
      const drawn = arcs(layoutDoughnut(doughnut([values]), BOX));
      expect(drawn).toHaveLength(values.length);
      const total = drawn.reduce((sum, one) => sum + sweep(one), 0);
      expect(Math.abs(total - 360)).toBeLessThan(0.01);
      drawn.slice(1).forEach((one, i) => {
        expect(one.start).toBe(drawn[i]!.end);
      });
      expect(drawn[0]!.start).toBe(-90);
      expect(drawn.at(-1)!.end).toBe(-90);
    }
  });

  it("writes every angle as the host reads it back: folded into (-180, 180]", () => {
    const drawn = arcs(layoutDoughnut(doughnut([[3, 7, 11, 13, 17, 19]]), BOX));
    for (const one of drawn) {
      for (const angle of [one.start, one.end]) {
        expect(angle).toBe(normalizeAngle(angle));
        expect(angle).toBeGreaterThan(-180);
        expect(angle).toBeLessThanOrEqual(180);
      }
    }
  });

  it("gives even the thinnest slice two different angles, never a full sweep", () => {
    const [rest, sliver] = arcs(
      layoutDoughnut(doughnut([[1_000_000, 1]]), BOX),
    );
    expect(rest!.start).not.toBe(rest!.end);
    expect(sliver!.start).not.toBe(sliver!.end);
  });

  it("draws nothing for a slice so thin it takes no angle the host could tell from none, and never a full sweep for it", () => {
    const out = arcs(layoutDoughnut(doughnut([[1e6, 1e-300, 1]]), BOX));
    expect(out.map((one) => one.name)).toEqual(["slice 0.0", "slice 0.2"]);
    for (const one of out) expect(one.start).not.toBe(one.end);
  });

  it("draws a ring whose other slices are thinner than any angle as the Donut it looks like, not as nothing", () => {
    // 1 of 10^12: 3.6e-10 of a degree, a thousandth of a micron at the rim.
    const out = layoutDoughnut(doughnut([[1e12, 1e-300, 1]]), BOX);
    expect(arcs(out)).toEqual([]);
    expect(donuts(out).map((one) => one.name)).toEqual(["slice 0.0"]);
  });

  it("draws a lone 100 % slice as one Donut, not an arc", () => {
    const out = layoutDoughnut(doughnut([[5, 0, 0]], { hole: 60 }), BOX);
    expect(arcs(out)).toEqual([]);
    expect(donuts(out)).toHaveLength(1);
    expect(donuts(out)[0]!.name).toBe("slice 0.0");
    expect(donuts(out)[0]!.color).toBe(DARK);
    expect(donuts(out)[0]!.thickness).toBeCloseTo(0.2, 12);
  });

  it("draws nothing for a zero, a negative and a blank slice, and keeps the colours' order", () => {
    const out = layoutDoughnut(doughnut([[3, 0, 2, null, -1, 5]]), BOX);
    expect(arcs(out).map((one) => one.name)).toEqual([
      "slice 0.0",
      "slice 0.2",
      "slice 0.5",
    ]);
    expect(arcs(out).map((one) => one.color)).toEqual([
      COLORS[0],
      COLORS[2],
      COLORS[5],
    ]);
    // The negative and the blank take no angle: 3 + 2 + 5 = 10.
    expect(arcs(out).map(sweep)).toEqual([108, 72, 180]);
  });

  it("draws no ring at all when no slice is positive, only what surrounds it", () => {
    const out = layoutDoughnut(doughnut([[0, -3, null]], { title: "T" }), BOX);
    expect(arcs(out)).toEqual([]);
    expect(donuts(out)).toEqual([]);
    expect(labels(out)).toEqual([]);
    expect(out.some((one) => one.name === "title")).toBe(true);
    expect(out.some((one) => one.name.startsWith("legend"))).toBe(true);
  });
});

describe("layoutDoughnut: the hole", () => {
  const plain = doughnut([[1, 2, 3]]);

  it("reads the hole as a share of the outer diameter: thickness is (1 - hole) / 2 of the box", () => {
    for (const [hole, thickness] of [
      [10, 0.45],
      [50, 0.25],
      [55, 0.225],
      [60, 0.2],
      [90, 0.05],
    ] as const) {
      const drawn = arcs(layoutDoughnut({ ...plain, hole }, BOX));
      for (const one of drawn) expect(one.thickness).toBeCloseTo(thickness, 12);
    }
  });

  it("leaves a hole whose diameter is that share of the doughnut's, centred", () => {
    const { cx, cy, radius } = circle(plain);
    const [first] = arcs(layoutDoughnut({ ...plain, hole: 40 }, BOX));
    const { inner, outer } = radii(first!);
    expect(outer).toBeCloseTo(radius, 9);
    expect(inner / outer).toBeCloseTo(0.4, 9);
    expect(first!.box.left + first!.box.width / 2).toBeCloseTo(cx, 9);
    expect(first!.box.top + first!.box.height / 2).toBeCloseTo(cy, 9);
    expect(first!.box.width).toBeCloseTo(first!.box.height, 9);
  });

  it("draws the hole Excel gives a new doughnut, 75 %, when the sheet did not say", () => {
    const [first] = arcs(layoutDoughnut(plain, BOX));
    expect(first!.thickness).toBeCloseTo(0.125, 12);
  });

  it("never gives a ring a thickness that would fill it solid", () => {
    for (const hole of [10, 33, 90]) {
      for (const rings of [1, 2, 5]) {
        const data = doughnut(
          Array.from({ length: rings }, () => [1, 2, 3]),
          { hole },
        );
        for (const one of arcs(layoutDoughnut(data, BOX))) {
          expect(one.thickness).toBeGreaterThan(0);
          expect(one.thickness).toBeLessThan(0.5);
        }
      }
    }
  });
});
