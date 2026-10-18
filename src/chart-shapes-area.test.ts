// The area kind as primitives: a rectangle and a triangle for each stretch
// between two points, the points at the centres of their category slots, the
// pieces overlapping by the seam, a value under zero hanging from the baseline,
// a blank cell drawn as a zero, several series painted first to last.
// Called on its own, with the plot and the scale worked out here from the
// same rules (the stacked kind is chart-shapes-area.stack.test.ts, the value
// labels chart-shapes-area.labels.test.ts).

import { describe, expect, it } from "vitest";
import {
  BOX,
  DARK,
  LIGHT,
  areaOf,
  fillRects,
  fills,
  frame,
  named,
  tipEnd,
  tipX,
  triangles,
} from "../test/chart-area.support";
import { layoutArea } from "./chart-shapes-area";
import { AREA_SEAM, slopePieces } from "./chart-shapes-area-pieces";

describe("layoutArea: an area", () => {
  const data = areaOf("area", [{ values: [10, 30, 20, 40] }]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 40 });

  it("is a rectangle and a triangle for each stretch between two points", () => {
    expect(fillRects(out).map((one) => one.name)).toEqual([
      "area 0.0",
      "area 0.1",
      "area 0.2",
    ]);
    expect(triangles(out).map((one) => one.name)).toEqual([
      "slope 0.0",
      "slope 0.1",
      "slope 0.2",
    ]);
  });

  it("tips each triangle at the higher end: right for a rise, left for a fall", () => {
    const [rise, fall, last] = triangles(out);
    expect(rise!.apex).toBeGreaterThan(0.9);
    expect(fall!.apex).toBeLessThan(0.1);
    // The end of the run reaches into nothing, so its tip is on the corner.
    expect(last!.apex).toBe(1);
    expect(triangles(out).every((one) => !one.down)).toBe(true);
  });

  it("puts the points at the centres of their category slots, where the labels sit", () => {
    // The tip of each triangle is at the higher point of its stretch.
    const [rise, fall, last] = triangles(out);
    expect(tipX(rise!)).toBeCloseTo(at.x(1));
    expect(tipX(fall!)).toBeCloseTo(at.x(1));
    expect(tipX(last!)).toBeCloseTo(at.x(3));
    // And the low end of the first one is at the first point.
    expect(rise!.box.left).toBeCloseTo(at.x(0));
  });

  it("stands each triangle on the lower end of its stretch and lifts it to the higher one", () => {
    const [rise, fall] = triangles(out);
    expect(rise!.box.top).toBeCloseTo(at.y(30));
    expect(rise!.box.top + rise!.box.height).toBeCloseTo(at.y(10));
    expect(fall!.box.top).toBeCloseTo(at.y(30));
    expect(fall!.box.top + fall!.box.height).toBeCloseTo(at.y(20));
  });

  it("is the pieces of each stretch, worked out from where its two points land", () => {
    const slope = (i: number, from: number, to: number) =>
      slopePieces(
        {
          left: at.x(i),
          right: at.x(i + 1),
          from: at.y(from),
          to: at.y(to),
          base: at.y(0),
        },
        {
          id: `0.${String(i)}`,
          color: LIGHT,
          seam: AREA_SEAM,
          first: i === 0,
          last: i === 2,
        },
      );
    expect(fills(out)).toEqual([
      ...slope(0, 10, 30),
      ...slope(1, 30, 20),
      ...slope(2, 20, 40),
    ]);
  });

  it("fills every piece with the series colour and draws no outline of its own", () => {
    expect(fills(out).every((one) => one.color === LIGHT)).toBe(true);
    // The zero line is the only line there is.
    expect(out.filter((one) => one.kind === "line")).toHaveLength(1);
  });

  it("rests the rectangles on the baseline, and draws the zero line over the fills", () => {
    const baseline = out.find((one) => one.name === "baseline")!;
    expect(baseline.box.top).toBeCloseTo(at.y(0));
    expect(out.indexOf(baseline)).toBeGreaterThan(
      out.indexOf(fills(out).at(-1)!),
    );
    for (const one of fillRects(out)) {
      expect(one.box.top + one.box.height).toBeCloseTo(at.y(0));
    }
  });

  it("names the categories under the plot, and has no legend for one series", () => {
    expect(named(out, "category ")).toHaveLength(4);
    expect(named(out, "legend")).toEqual([]);
  });

  it("draws the title first", () => {
    const titled = layoutArea({ ...data, title: "Net debt" }, BOX);
    expect(titled[0]).toMatchObject({ name: "title", text: "Net debt" });
  });
});

describe("layoutArea: the seam", () => {
  const data = areaOf("area", [{ values: [10, 30, 20, 40] }]);
  const out = layoutArea(data, BOX);
  const slot = frame(data, { min: 0, max: 40 }).slot;

  it("lets a rectangle reach across into the next stretch, except at the end of a run", () => {
    const widths = fillRects(out).map((one) => one.box.width);
    expect(widths[0]).toBeCloseTo(slot + AREA_SEAM);
    expect(widths[1]).toBeCloseTo(slot + AREA_SEAM);
    expect(widths[2]).toBeCloseTo(slot);
  });

  it("lets it reach up into its triangle", () => {
    const [rise] = triangles(out);
    const [first] = fillRects(out);
    expect(first!.box.top).toBeCloseTo(
      rise!.box.top + rise!.box.height - AREA_SEAM,
    );
  });
});

describe("layoutArea: a value under zero", () => {
  const data = areaOf("area", [{ values: [10, -10, 10] }]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: -10, max: 10 });

  it("splits a stretch through zero into a triangle on each side of the baseline", () => {
    expect(
      triangles(out).map((one) => [one.name, tipEnd(one), one.down]),
    ).toEqual([
      ["slope 0.0.0", "left", false],
      ["slope 0.0.1", "right", true],
      ["slope 0.1.0", "left", true],
      ["slope 0.1.1", "right", false],
    ]);
    expect(fillRects(out)).toEqual([]);
  });

  it("hangs the lower triangles from the baseline", () => {
    const hanging = triangles(out).filter((one) => one.down);
    for (const one of hanging) expect(one.box.top).toBeCloseTo(at.y(0));
    expect(hanging[0]!.box.top + hanging[0]!.box.height).toBeCloseTo(at.y(-10));
  });

  it("puts the zero line between the deepest value and the highest", () => {
    const baseline = out.find((one) => one.name === "baseline")!;
    expect(baseline.box.top).toBeCloseTo(at.y(0));
    expect(at.y(0)).toBeGreaterThan(at.plot.top);
    expect(at.y(0)).toBeLessThan(at.plot.top + at.plot.height);
  });

  it("fills an area that stays under zero downwards from the baseline", () => {
    const below = areaOf("area", [{ values: [-10, -30, -20] }]);
    const drawn = layoutArea(below, BOX);
    const line = frame(below, { min: -30, max: 0 });
    for (const one of fillRects(drawn)) {
      expect(one.box.top).toBeCloseTo(line.y(0));
    }
    expect(triangles(drawn).every((one) => one.down)).toBe(true);
    expect(triangles(drawn).map(tipEnd)).toEqual(["right", "left"]);
  });
});

// Excel cannot leave a gap in an area chart: it plots a blank cell as a zero.
describe("layoutArea: a blank cell", () => {
  it("is a zero: the area dips to the baseline there and does not break", () => {
    const data = areaOf("area", [{ values: [10, null, 30, 40, null, 5] }]);
    const out = layoutArea(data, BOX);
    const at = frame(data, { min: 0, max: 40 });
    // Every stretch is drawn; a stretch that ends on the baseline is only a
    // triangle, so the rectangle is the one between 30 and 40.
    expect(fillRects(out).map((one) => one.name)).toEqual(["area 0.2"]);
    const drawn = triangles(out);
    expect(drawn.map((one) => one.name)).toEqual([
      "slope 0.0",
      "slope 0.1",
      "slope 0.2",
      "slope 0.3",
      "slope 0.4",
    ]);
    const [down, up] = drawn;
    expect(down!.box.top).toBeCloseTo(at.y(10));
    expect(down!.box.top + down!.box.height).toBeCloseTo(at.y(0));
    expect(tipEnd(down!)).toBe("left");
    expect(up!.box.top).toBeCloseTo(at.y(30));
    expect(tipEnd(up!)).toBe("right");
  });

  it("draws a series of only blanks as a flat zero: no fill, no label, still the zero line", () => {
    const data = areaOf("area", [{ values: [null, null, null] }]);
    const out = layoutArea(data, BOX);
    expect(fills(out)).toEqual([]);
    expect(named(out, "label ")).toEqual([]);
    expect(out.some((one) => one.name === "baseline")).toBe(true);
  });

  it("keeps the run going over the dip, only the last stretch not reaching past itself", () => {
    const data = areaOf("area", [{ values: [10, 20, null, 30, 40, 50] }]);
    const out = layoutArea(data, BOX);
    const bars = fillRects(out);
    expect(bars.map((one) => one.name)).toEqual([
      "area 0.0",
      "area 0.3",
      "area 0.4",
    ]);
    const slot = frame(data, { min: 0, max: 50 }).slot;
    expect(bars.map((one) => one.box.width)).toEqual([
      expect.closeTo(slot + AREA_SEAM, 6),
      expect.closeTo(slot + AREA_SEAM, 6),
      expect.closeTo(slot, 6),
    ]);
  });
});

describe("layoutArea: several series that are not stacked", () => {
  const data = areaOf("area", [
    { values: [30, 30, 30], color: DARK },
    { values: [10, 20, 10], color: LIGHT },
  ]);
  const out = layoutArea(data, BOX);

  it("paints the first series first, so the later ones lie over it", () => {
    const names = fills(out).map((one) => one.name);
    expect(names.indexOf("area 0.0")).toBeLessThan(names.indexOf("area 1.0"));
    expect(names.indexOf("area 0.1")).toBeLessThan(names.indexOf("area 1.1"));
  });

  it("colours each series in its own colour", () => {
    const colorOf = (prefix: string) =>
      fills(out).find((one) => one.name.startsWith(prefix))!.color;
    expect([colorOf("area 0."), colorOf("area 1.")]).toEqual([DARK, LIGHT]);
  });

  it("has a legend for them, and every shape its own name", () => {
    expect(named(out, "legend swatch")).toHaveLength(2);
    const names = out.map((one) => one.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe("layoutArea: a single point", () => {
  it("fills nothing, and still draws its label and the zero line", () => {
    const data = areaOf("area", [{ values: [10] }]);
    const out = layoutArea(data, BOX);
    expect(fills(out)).toEqual([]);
    expect(named(out, "label ")).toHaveLength(1);
    expect(out.some((one) => one.name === "baseline")).toBe(true);
  });
});
