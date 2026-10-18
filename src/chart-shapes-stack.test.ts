// Stacked columns: one rectangle per segment per category, positives up from
// the baseline and negatives below it, a blank cell leaving its slot empty, a
// segment label centred inside when it fits. 100 % stacks, totals, stacked
// bars and the caps are in chart-shapes-stack.percent.test.ts.

import { describe, expect, it } from "vitest";
import { layoutStack } from "./chart-shapes-stack";
import type { Primitive, Rect, Text } from "./chart-shapes";
import type { ChartData, ChartSeries } from "./link/chart-model";

const BOX = { left: 100, top: 100, width: 480, height: 270 };
const base = {
  v: 1 as const,
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
  title: null,
};
const DARK = "#14213D";
const LIGHT = "#F2D16B";

function series(
  name: string,
  values: (number | null)[],
  color = DARK,
  extra: Partial<ChartSeries> = {},
): ChartSeries {
  return {
    name,
    values,
    labels: values.map((value) => (value === null ? "" : String(value))),
    colors: values.map(() => color),
    ...extra,
  };
}

function stack(
  kind: "stackedColumn" | "stackedBar",
  list: ChartSeries[],
  extra: Partial<ChartData> = {},
): ChartData {
  const points = list[0]!.values.length;
  return {
    ...base,
    kind,
    categories: Array.from({ length: points }, (_one, i) => `C${String(i)}`),
    series: list,
    ...extra,
  };
}

const bars = (out: Primitive[]) =>
  out.filter(
    (one): one is Rect => one.kind === "rect" && one.name.startsWith("bar "),
  );
const texts = (out: Primitive[], prefix: string) =>
  out.filter(
    (one): one is Text => one.kind === "text" && one.name.startsWith(prefix),
  );
const named = (out: Primitive[], name: string) =>
  out.find((one) => one.name === name)!;
const baseline = (out: Primitive[]) => named(out, "baseline");

function inside(one: Primitive): boolean {
  const b = one.box;
  return (
    Number.isFinite(b.left + b.top + b.width + b.height) &&
    b.width >= 0 &&
    b.height >= 0 &&
    b.left >= BOX.left - 0.01 &&
    b.top >= BOX.top - 0.01 &&
    b.left + b.width <= BOX.left + BOX.width + 0.01 &&
    b.top + b.height <= BOX.top + BOX.height + 0.01
  );
}

const TWO = [series("a", [60, 40, 80]), series("b", [40, 60, 20], LIGHT)];

describe("a stacked column chart", () => {
  const out = layoutStack(stack("stackedColumn", TWO), BOX);

  it("draws one rectangle per segment, series by series, with a column per category", () => {
    const drawn = bars(out);
    expect(drawn.map((one) => one.name)).toEqual([
      "bar 0.0",
      "bar 0.1",
      "bar 0.2",
      "bar 1.0",
      "bar 1.1",
      "bar 1.2",
    ]);
    // The two segments of a category share one column.
    expect(drawn[3]!.box.left).toBeCloseTo(drawn[0]!.box.left);
    expect(drawn[3]!.box.width).toBeCloseTo(drawn[0]!.box.width);
  });

  it("piles the second segment exactly on the first, in proportion to the values", () => {
    const [a0, , , b0] = bars(out);
    expect(b0!.box.top + b0!.box.height).toBeCloseTo(a0!.box.top);
    expect(a0!.box.height / b0!.box.height).toBeCloseTo(60 / 40);
  });

  it("reaches the top of the plot with the tallest stack and rests on the baseline", () => {
    const drawn = bars(out);
    const tops = drawn.map((one) => one.box.top);
    const bottom = baseline(out).box.top;
    // C0 and C1 and C2 all total 100: every column spans the same height.
    expect(Math.min(...tops)).toBeLessThan(bottom);
    const [a0, a1] = [drawn[0]!, drawn[1]!];
    expect(a0!.box.top + a0!.box.height).toBeCloseTo(bottom);
    expect(a1!.box.top + a1!.box.height).toBeCloseTo(bottom);
  });

  it("labels each segment inside it, white on a dark fill and the ink on a light one", () => {
    expect(named(out, "label 0.0")).toMatchObject({
      text: "60",
      color: "#FFFFFF",
      align: "c",
    });
    expect(named(out, "label 1.0")).toMatchObject({
      text: "40",
      color: "#333333",
    });
  });

  it("names the categories under the plot and the series in a legend", () => {
    expect(texts(out, "category ").map((one) => one.text)).toEqual([
      "C0",
      "C1",
      "C2",
    ]);
    expect(texts(out, "legend ").map((one) => one.text)).toEqual(["a", "b"]);
  });

  it("keeps every primitive inside the box and gives each its own name", () => {
    expect(out.every(inside)).toBe(true);
    expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
  });
});

describe("negative values", () => {
  const data = stack("stackedColumn", [
    series("up", [50, 30]),
    series("down", [-20, -40], LIGHT),
  ]);
  const out = layoutStack(data, BOX);

  it("hang below the baseline and stack downwards", () => {
    const [up0, , down0] = bars(out);
    const zero = baseline(out).box.top;
    expect(up0!.box.top + up0!.box.height).toBeCloseTo(zero);
    expect(down0!.box.top).toBeCloseTo(zero);
    expect(down0!.box.height).toBeGreaterThan(0);
  });

  it("make room for the deepest column below the baseline", () => {
    const deepest = bars(out)[3]!;
    expect(deepest.box.top + deepest.box.height).toBeLessThanOrEqual(
      BOX.top + BOX.height,
    );
  });

  it("stack a second negative segment below the first", () => {
    const two = layoutStack(
      stack("stackedColumn", [
        series("a", [-10, -10]),
        series("b", [-30, -30], LIGHT),
      ]),
      BOX,
    );
    const [a0, , b0] = bars(two);
    expect(b0!.box.top).toBeCloseTo(a0!.box.top + a0!.box.height);
  });
});

describe("a blank cell", () => {
  const data = stack("stackedColumn", [
    series("a", [60, null, 80]),
    series("b", [40, 60, 20], LIGHT),
  ]);
  const out = layoutStack(data, BOX);

  it("draws no segment and no label in its slot, and nothing else moves", () => {
    expect(bars(out).map((one) => one.name)).toEqual([
      "bar 0.0",
      "bar 0.2",
      "bar 1.0",
      "bar 1.1",
      "bar 1.2",
    ]);
    expect(texts(out, "label ").map((one) => one.name)).not.toContain(
      "label 0.1",
    );
    // The segment above the gap rests on the baseline: the gap is zero tall.
    expect(
      named(out, "bar 1.1").box.top + named(out, "bar 1.1").box.height,
    ).toBeCloseTo(baseline(out).box.top);
  });

  it("draws a category of nothing but gaps as an empty slot with its name", () => {
    const empty = layoutStack(
      stack("stackedColumn", [
        series("a", [1, null]),
        series("b", [2, null], LIGHT),
      ]),
      BOX,
    );
    expect(bars(empty)).toHaveLength(2);
    expect(texts(empty, "category ").map((one) => one.text)).toEqual([
      "C0",
      "C1",
    ]);
    expect(empty.every(inside)).toBe(true);
  });
});

describe("segment labels", () => {
  it("drops the label of a segment too thin to hold it, and keeps the segment", () => {
    const out = layoutStack(
      stack("stackedColumn", [
        series("big", [1000, 1000]),
        series("sliver", [1, 1], LIGHT),
      ]),
      BOX,
    );
    expect(bars(out)).toHaveLength(4);
    expect(texts(out, "label ").map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.1",
    ]);
  });

  it("draws none for a series without data labels, and still its segments", () => {
    const out = layoutStack(
      stack("stackedColumn", [
        series("a", [60, 40], DARK, { dataLabels: false }),
        series("b", [40, 60], LIGHT),
      ]),
      BOX,
    );
    expect(texts(out, "label ").map((one) => one.name)).toEqual([
      "label 1.0",
      "label 1.1",
    ]);
    expect(bars(out)).toHaveLength(4);
  });

  it("draws none at all when the columns are too narrow for the widest label", () => {
    const many = Array.from({ length: 120 }, (_one, i) => 1000 + i);
    const out = layoutStack(
      stack("stackedColumn", [series("a", many), series("b", many, LIGHT)]),
      BOX,
    );
    expect(texts(out, "label ")).toEqual([]);
    expect(bars(out)).toHaveLength(240);
  });
});
