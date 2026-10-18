// The stacked area as primitives: painter's order (each series drawn from the
// baseline up to its cumulative top, the highest series first, so the lower
// ones paint over it), a blank cell adding nothing to the stack (a band of no
// height, as Excel plots it), a 100 % stack as each point's shares, and the
// value label centred in the band of its series.

import { describe, expect, it } from "vitest";
import { LABEL_HEIGHT, MIN_SEGMENT, type Text } from "./chart-shapes";
import {
  BOX,
  DARK,
  INK,
  LIGHT,
  WHITE,
  areaOf,
  fillRects,
  fills,
  frame,
  named,
  tipEnd,
  triangles,
} from "../test/chart-area.support";
import { layoutArea } from "./chart-shapes-area";
import { AREA_SEAM } from "./chart-shapes-area-pieces";

describe("layoutArea: a stacked area", () => {
  const data = areaOf("stackedArea", [
    { values: [10, 20, 30], color: DARK },
    { values: [5, 5, 5], color: LIGHT },
  ]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 35 });

  it("paints the highest series first and the lowest last, so it lies over the rest", () => {
    const names = fills(out).map((one) => one.name);
    expect(names.indexOf("area 1.0")).toBeLessThan(names.indexOf("area 0.0"));
    expect(names.indexOf("slope 1.1")).toBeLessThan(names.indexOf("slope 0.1"));
    // And all of the upper series before any of the lower.
    const lastUpper = Math.max(
      ...names.flatMap((name, i) => (name.includes(" 1.") ? [i] : [])),
    );
    const firstLower = Math.min(
      ...names.flatMap((name, i) => (name.includes(" 0.") ? [i] : [])),
    );
    expect(lastUpper).toBeLessThan(firstLower);
  });

  it("draws each series from the baseline up to its cumulative top", () => {
    const upper = fillRects(out).filter((one) =>
      one.name.startsWith("area 1."),
    );
    // 10 + 5, 20 + 5 and 30 + 5 at the points; the rectangle is the lower end.
    expect(upper[0]!.box.top).toBeCloseTo(at.y(15) - AREA_SEAM);
    expect(upper[0]!.box.top + upper[0]!.box.height).toBeCloseTo(at.y(0));
    const lower = fillRects(out).filter((one) =>
      one.name.startsWith("area 0."),
    );
    expect(lower[0]!.box.top).toBeCloseTo(at.y(10) - AREA_SEAM);
  });

  it("scales over the highest cumulative top", () => {
    const top = Math.min(...fills(out).map((one) => one.box.top));
    expect(top).toBeCloseTo(at.plot.top);
  });

  it("colours each series in its own colour", () => {
    const colorOf = (prefix: string) =>
      fills(out).find((one) => one.name.startsWith(prefix))!.color;
    expect([colorOf("area 0."), colorOf("area 1.")]).toEqual([DARK, LIGHT]);
  });

  it("has a legend with a swatch per series and every shape its own name", () => {
    expect(named(out, "legend swatch")).toHaveLength(2);
    const names = out.map((one) => one.name);
    expect(new Set(names).size).toBe(names.length);
  });
});

describe("layoutArea: a blank cell in a stack", () => {
  const data = areaOf("stackedArea", [
    { values: [10, null, 30], color: DARK },
    { values: [5, 5, 5], color: LIGHT },
  ]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 35 });

  it("keeps a band of no height: the series dips to the stack below it and both series keep every stretch", () => {
    const names = fills(out).map((one) => one.name);
    expect(names.filter((name) => name.includes(" 0."))).toEqual([
      "slope 0.0",
      "slope 0.1",
    ]);
    expect(names.filter((name) => name.startsWith("area 1."))).toEqual([
      "area 1.0",
      "area 1.1",
    ]);
  });

  it("adds nothing to the stack there: the upper series sits on 15 at the first point and on 5 over the blank", () => {
    const first = triangles(out).find((one) => one.name === "slope 1.0")!;
    expect(first.box.top).toBeCloseTo(at.y(15));
    expect(first.box.top + first.box.height).toBeCloseTo(at.y(5));
  });

  it("draws the blank series right down to the baseline and back up", () => {
    const [down, up] = triangles(out).filter((one) => one.name.includes(" 0."));
    expect(down!.box.top).toBeCloseTo(at.y(10));
    expect(down!.box.top + down!.box.height).toBeCloseTo(at.y(0));
    expect(up!.box.top).toBeCloseTo(at.y(30));
    expect(up!.box.top + up!.box.height).toBeCloseTo(at.y(0));
  });
});

// The reviewer's failing inputs: a series that starts with blanks, and one
// with a blank in the middle, each under a series that must stay in its own
// colour above the first one's line.
describe("layoutArea: a blank cell under another series", () => {
  it("starts a series late: its fill climbs from the baseline between the last blank and the first value", () => {
    const data = areaOf("stackedArea", [
      { values: [null, null, 5, 6, 7, 8], color: DARK },
      { values: [3, 3, 3, 3, 3, 3], color: LIGHT },
    ]);
    const out = layoutArea(data, BOX);
    const at = frame(data, { min: 0, max: 11 });
    const climb = triangles(out).find((one) => one.name === "slope 0.1")!;
    expect(climb.box.top).toBeCloseTo(at.y(5));
    expect(climb.box.top + climb.box.height).toBeCloseTo(at.y(0));
    expect(tipEnd(climb)).toBe("right");
    // Nothing of the first series before it: two flat stretches on the baseline.
    const names = fills(out).map((one) => one.name);
    expect(names.filter((name) => name.endsWith(" 0.0"))).toEqual([]);
  });

  it("dips a series in the middle: both of its stretches around the blank are drawn", () => {
    const data = areaOf("stackedArea", [
      { values: [5, 5, null, 5, 5], color: DARK },
      { values: [3, 3, 3, 3, 3], color: LIGHT },
    ]);
    const out = layoutArea(data, BOX);
    const slopes = triangles(out)
      .filter((one) => one.name.includes(" 0."))
      .map((one) => one.name);
    expect(slopes).toEqual(["slope 0.1", "slope 0.2"]);
  });
});

describe("layoutArea: a 100 % stacked area", () => {
  const data = areaOf(
    "stackedArea",
    [
      { values: [60, 10, 5], color: DARK },
      { values: [40, 30, 15], color: LIGHT },
    ],
    { v: 2, percent: true },
  );
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 1 });

  it("fills the plot from the baseline to the top at every point", () => {
    const upper = fillRects(out).filter((one) =>
      one.name.startsWith("area 1."),
    );
    expect(upper).toHaveLength(2);
    for (const one of upper) expect(one.box.top).toBeCloseTo(at.plot.top);
    // The top series is flat at one hundred per cent: no triangle above it.
    expect(triangles(out).filter((one) => one.name.includes(" 1."))).toEqual(
      [],
    );
  });

  it("draws the lower series as its share of each point", () => {
    const [first] = triangles(out).filter((one) => one.name.includes(" 0."));
    // 60 % at the first point, 20 % at the second (10 of 40).
    expect(first!.box.top).toBeCloseTo(at.y(0.6));
    expect(first!.box.top + first!.box.height).toBeCloseTo(at.y(0.25));
  });

  it("shares a point of mixed signs over its absolute total, a point of zeros stays flat", () => {
    const mixed = areaOf(
      "stackedArea",
      [
        { values: [30, 0], color: DARK },
        { values: [-10, 0], color: LIGHT },
      ],
      { v: 2, percent: true },
    );
    const drawn = layoutArea(mixed, BOX);
    // 30 of |30| + |-10| is 75 %, and the second series adds -25 on top of it.
    const line = frame(mixed, { min: 0, max: 0.75 });
    const lower = triangles(drawn).find((one) =>
      one.name.startsWith("slope 0."),
    );
    expect(lower!.box.top).toBeCloseTo(line.y(0.75));
  });
});

describe("layoutArea: the labels of a stacked area", () => {
  const data = areaOf("stackedArea", [
    { values: [100, 100, 100], color: DARK },
    { values: [100, 100, 100], color: LIGHT },
  ]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 200 });
  const shown = named(out, "label ") as Text[];

  it("centres each value in the band of its series", () => {
    expect(shown).toHaveLength(6);
    const lower = shown.find((one) => one.name === "label 0.1")!;
    expect(lower.box.top + lower.box.height / 2).toBeCloseTo(
      (at.y(0) + at.y(100)) / 2,
    );
    const upper = shown.find((one) => one.name === "label 1.1")!;
    expect(upper.box.top + upper.box.height / 2).toBeCloseTo(
      (at.y(100) + at.y(200)) / 2,
    );
  });

  it("writes white on a dark band and ink on a light one", () => {
    expect(shown.find((one) => one.name === "label 0.1")!.color).toBe(WHITE);
    expect(shown.find((one) => one.name === "label 1.1")!.color).toBe(INK);
  });

  it("draws the labels over the fills", () => {
    const firstLabel = out.indexOf(shown[0]!);
    expect(firstLabel).toBeGreaterThan(out.indexOf(fills(out).at(-1)!));
  });

  it("drops the label of a band too thin to hold one", () => {
    const thin = areaOf("stackedArea", [
      { values: [1000, 1000, 1000], color: DARK },
      { values: [1, 1, 1], color: LIGHT },
    ]);
    const drawn = named(layoutArea(thin, BOX), "label ");
    expect(drawn.map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.1",
      "label 0.2",
    ]);
    expect(MIN_SEGMENT).toBeGreaterThan(1);
    expect(LABEL_HEIGHT).toBeGreaterThanOrEqual(MIN_SEGMENT);
  });

  it("labels only the last point of a series without data labels, in its band", () => {
    const quiet = areaOf("stackedArea", [
      { values: [100, 100, 100], color: DARK, extra: { dataLabels: false } },
      { values: [100, 100, 100], color: LIGHT, extra: { dataLabels: false } },
    ]);
    const drawn = named(layoutArea(quiet, BOX), "label ");
    expect(drawn.map((one) => one.name)).toEqual(["label 0.2", "label 1.2"]);
  });
});

describe("layoutArea: a dense stack", () => {
  // 12 series of 120 points stepping a unit up and down on a scale of a
  // thousand: the lower layers rise a fraction of a point a stretch, which their
  // rectangle already covers.
  const data = areaOf(
    "stackedArea",
    Array.from({ length: 12 }, (_one, s) => ({
      values: Array.from({ length: 120 }, (_two, i) => 100 + (i % 2)),
      color: s % 2 === 0 ? DARK : LIGHT,
      extra: { dataLabels: false },
    })),
  );
  const out = layoutArea(data, BOX);

  it("emits no triangle no taller than the seam", () => {
    const standing = triangles(out);
    expect(standing.length).toBeGreaterThan(0);
    expect(standing.every((one) => one.box.height > AREA_SEAM)).toBe(true);
  });

  it("draws the gentle stretches as their rectangle alone", () => {
    expect(fillRects(out)).toHaveLength(12 * 119);
    // At least the lowest layer's 119 stretches have no triangle.
    expect(triangles(out).length).toBeLessThan(12 * 119 - 119);
  });
});

describe("layoutArea: a stacked segment that starts late", () => {
  // The first series is blank until the third point: it adds nothing to the
  // stack there, then rises from the baseline in its own colour while the series
  // above keeps its own over the same points.
  const data = areaOf("stackedArea", [
    { values: [null, null, 5, 6, 7, 8], color: DARK },
    { values: [3, 3, 3, 3, 3, 3], color: LIGHT },
  ]);
  const out = layoutArea(data, BOX);
  const at = frame(data, { min: 0, max: 11 });

  it("tapers in its own colour from the point it starts at", () => {
    expect(named(out, "slope 0.0")).toEqual([]);
    const [taper, ...rest] = triangles(named(out, "slope 0.1"));
    expect(rest).toEqual([]);
    expect(taper!.color).toBe(DARK);
    expect(tipEnd(taper!)).toBe("right");
    expect(taper!.box.top).toBeCloseTo(at.y(5));
    expect(taper!.box.top + taper!.box.height).toBeCloseTo(at.y(0));
  });

  it("leaves the series above in its own colour all along", () => {
    const upper = fills(out).filter((one) => one.name.includes(" 1."));
    expect(upper.length).toBeGreaterThan(0);
    expect(upper.every((one) => one.color === LIGHT)).toBe(true);
    const lower = fills(out).filter((one) => one.name.includes(" 0."));
    expect(lower.every((one) => one.color === DARK)).toBe(true);
  });
});
