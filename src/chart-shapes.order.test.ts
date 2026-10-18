// Where the categories land and what a blank cell or a missing data label
// draws, through the one entry layoutChart: a bar chart's first category sits
// at the bottom as in Excel only when the sheet says its axis is not reversed
// (an explicit false) and on top when the flag is absent or true, a column chart
// runs last to first when reversed, a clustered bar keeps Excel's series order,
// a waterfall step with no value writes no empty text, a pie without data
// labels draws no slice label.

import { describe, expect, it } from "vitest";
import {
  layoutChart,
  type Primitive,
  type Rect,
  type Text,
  type Wedge,
} from "./chart-shapes";
import type { ChartData } from "./link/chart-model";

const box = { left: 100, top: 100, width: 480, height: 270 };
const base = {
  v: 1 as const,
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
  title: "Revenue",
};
const rects = (p: Primitive[]) =>
  p.filter((x): x is Rect => x.kind === "rect" && x.name.startsWith("bar"));
const texts = (p: Primitive[]) => p.filter((x): x is Text => x.kind === "text");
const inside = (b: {
  left: number;
  top: number;
  width: number;
  height: number;
}) =>
  b.width >= 0 &&
  b.height >= 0 &&
  b.left >= box.left - 0.01 &&
  b.top >= box.top - 0.01 &&
  b.left + b.width <= box.left + box.width + 0.01 &&
  b.top + b.height <= box.top + box.height + 0.01;

const column: ChartData = {
  ...base,
  kind: "column",
  categories: ["A", "B", "C"],
  series: [
    {
      name: "s",
      values: [100, -50, 200],
      labels: ["100", "(50)", "200"],
      colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    },
  ],
};

describe("layoutChart: the order of the categories", () => {
  // Excel draws a bar chart's first category at the bottom, and its last on
  // top, unless the category axis is reversed. The payload says so with an
  // explicit false; a payload from before the flag has none and is drawn the
  // way every pane drew it, first category on top.
  it("puts a bar chart's first category at the bottom when the sheet says its axis is not reversed", () => {
    const out = layoutChart(
      { ...column, kind: "bar", categoriesReversed: false },
      box,
    );
    const labels = out.filter(
      (x): x is Text => x.kind === "text" && x.name.startsWith("category"),
    );
    const topOf = (name: string) =>
      labels.find((x) => x.text === name)!.box.top;
    expect(topOf("A")).toBeGreaterThan(topOf("B"));
    expect(topOf("B")).toBeGreaterThan(topOf("C"));
    // The bars follow their names: names count from the top as drawn, so the
    // first bar is category C, the longest.
    const bar = (j: number) =>
      out.find((x) => x.name === `bar 0.${String(j)}`)!;
    expect(bar(0).box.width).toBeGreaterThan(bar(2).box.width);
    expect(bar(0).box.top).toBeLessThan(bar(2).box.top);
  });
  it("puts a bar chart's first category on top when its axis is reversed or the flag is absent", () => {
    for (const extra of [{ categoriesReversed: true }, {}]) {
      const out = layoutChart({ ...column, kind: "bar", ...extra }, box);
      const labels = out.filter(
        (x): x is Text => x.kind === "text" && x.name.startsWith("category"),
      );
      const topOf = (name: string) =>
        labels.find((x) => x.text === name)!.box.top;
      expect(topOf("A")).toBeLessThan(topOf("B"));
      expect(topOf("B")).toBeLessThan(topOf("C"));
      const [a, , c] = rects(out);
      expect(a!.box.top).toBeLessThan(c!.box.top);
      expect(a!.box.width).toBeLessThan(c!.box.width);
    }
  });
  it("runs a column chart first to last, and last to first once its axis is reversed", () => {
    const left = (data: ChartData, name: string) =>
      layoutChart(data, box)
        .filter(
          (x): x is Text => x.kind === "text" && x.name.startsWith("category"),
        )
        .find((x) => x.text === name)!.box.left;
    expect(left(column, "A")).toBeLessThan(left(column, "C"));
    const reversed = { ...column, categoriesReversed: true };
    expect(left(reversed, "A")).toBeGreaterThan(left(reversed, "C"));
  });
  it("keeps a clustered bar's series in Excel's order: the first series lowest in its cluster", () => {
    const two: ChartData = {
      ...column,
      kind: "bar",
      series: [
        { ...column.series[0]!, name: "first" },
        { ...column.series[0]!, name: "second" },
      ],
    };
    const rowTop = (data: ChartData, name: string) =>
      layoutChart(data, box).find((x) => x.name === name)!.box.top;
    // An axis the sheet says is not reversed: display index 0 is the last
    // category; its two bars: series 1 above series 0.
    const upright = { ...two, categoriesReversed: false };
    expect(rowTop(upright, "bar 1.0")).toBeLessThan(rowTop(upright, "bar 0.0"));
    // Reversed, or a payload from before the flag, the first series leads:
    // series 0 above series 1.
    for (const extra of [{ categoriesReversed: true }, {}]) {
      const topDown = { ...two, ...extra };
      expect(rowTop(topDown, "bar 0.0")).toBeLessThan(
        rowTop(topDown, "bar 1.0"),
      );
    }
  });
});

describe("layoutChart waterfall with a blank step", () => {
  it("draws the step as nothing and writes no empty text box for it", () => {
    const bridge: ChartData = {
      ...column,
      kind: "waterfall",
      categories: ["Open", "Blank", "Down", "Close"],
      series: [
        {
          name: "s",
          values: [100, null, -30, 70],
          labels: ["100", "", "(30)", "70"],
          colors: ["#14213D", "#2EC4B6", "#B27E54", "#14213D"],
        },
      ],
    };
    const out = layoutChart(bridge, box);
    expect(texts(out).some((t) => t.name === "label 0.1")).toBe(false);
    expect(texts(out).filter((t) => t.name.startsWith("label"))).toHaveLength(
      3,
    );
    expect(texts(out).every((t) => t.text !== "")).toBe(true);
    expect(rects(out)).toHaveLength(4);
    expect(rects(out).every((r) => inside(r.box))).toBe(true);
  });
});

describe("layoutChart pie without data labels", () => {
  const pie: ChartData = {
    ...column,
    kind: "pie",
    categories: ["A", "B", "C", "D"],
    series: [
      {
        name: "s",
        values: [1, 1, 2, 0],
        labels: ["1", "1", "2", "0"],
        colors: ["#14213D", "#2EC4B6", "#B27E54", "#000000"],
      },
    ],
  };
  it("draws the slices but no slice label when the series has no data labels", () => {
    const quiet: ChartData = {
      ...pie,
      series: [{ ...pie.series[0]!, dataLabels: false }],
    };
    const out = layoutChart(quiet, box);
    expect(out.filter((x): x is Wedge => x.kind === "wedge")).toHaveLength(3);
    expect(texts(out).filter((t) => t.name.startsWith("label"))).toEqual([]);
    expect(texts(out).some((t) => t.name.startsWith("legend"))).toBe(true);
  });
});
