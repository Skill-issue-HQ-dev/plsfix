// Value labels on horizontal bars. A bar chart with many rows or many series
// per row drops its value labels once a bar's share of its row is shorter than
// a label is high, the way a column chart drops them when its slot is narrower
// than the widest label: 40 rows of 6 series or 120 rows of 3 used to print
// thousands of labels over each other. Roomy charts keep every label.

import { describe, expect, it } from "vitest";
import {
  BAR_FILL,
  LABEL_HEIGHT,
  layoutChart,
  type Primitive,
} from "./chart-shapes";
import { overlaps } from "./layout";
import type { ChartData, ChartKind } from "./link/chart-model";

const BOX = { left: 10, top: 20, width: 540, height: 324 };
// Skip a row count that lands exactly on the threshold: the last bit of a
// division is not what the rule is about.
const EDGE = 1e-6;

function chartOf(
  kind: ChartKind,
  rows: number,
  seriesCount: number,
  extra: Partial<ChartData> = {},
): ChartData {
  const categories = Array.from({ length: rows }, (_one, i) => `C${String(i)}`);
  return {
    v: 1,
    kind,
    title: "T",
    categories,
    series: Array.from({ length: seriesCount }, (_one, j) => ({
      name: `S${String(j)}`,
      values: categories.map((_c, i) => (i + 1) * (j + 1)),
      labels: categories.map((_c, i) => String((i + 1) * (j + 1))),
      colors: categories.map(() => "#2EC4B6"),
    })),
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
    categoriesReversed: true,
    ...extra,
  };
}

const valueLabels = (out: Primitive[]) =>
  out.filter((one) => one.kind === "text" && one.name.startsWith("label "));
const bars = (out: Primitive[]) =>
  out.filter((one) => one.kind === "rect" && one.name.startsWith("bar "));

// A row's share of the plot, read back off the bars the layout drew: a bar is
// BAR_FILL of its row's share, so the share is the bar's height over BAR_FILL.
function rowShare(out: Primitive[]): number {
  return bars(out)[0]!.box.height / BAR_FILL;
}

function noLabelOverlaps(out: Primitive[]): boolean {
  const labels = valueLabels(out);
  return labels.every((one, i) =>
    labels.slice(i + 1).every((other) => !overlaps(one.box, other.box)),
  );
}

describe("a clustered bar chart with too little room for its value labels", () => {
  it("draws every bar and no label at 40 rows of 6 series", () => {
    const out = layoutChart(chartOf("bar", 40, 6), BOX);
    expect(bars(out)).toHaveLength(240);
    expect(valueLabels(out)).toHaveLength(0);
    expect(noLabelOverlaps(out)).toBe(true);
  });

  it("draws every bar and no label at 120 rows of 3 series", () => {
    const out = layoutChart(chartOf("bar", 120, 3), BOX);
    expect(bars(out)).toHaveLength(360);
    expect(valueLabels(out)).toHaveLength(0);
    expect(noLabelOverlaps(out)).toBe(true);
  });

  it("still has its category names, thinned as before", () => {
    const out = layoutChart(chartOf("bar", 120, 3), BOX);
    const names = out.filter(
      (one) => one.kind === "text" && one.name.startsWith("category "),
    );
    expect(names.length).toBeGreaterThan(0);
    expect(names.length).toBeLessThan(120);
  });

  it("keeps every label when the bars have room, all of them clear of each other", () => {
    const out = layoutChart(chartOf("bar", 4, 2), BOX);
    expect(valueLabels(out)).toHaveLength(8);
    expect(noLabelOverlaps(out)).toBe(true);
  });

  it("draws labels exactly while a bar's share of its row is a label high", () => {
    for (const seriesCount of [1, 2, 3, 6]) {
      for (let rows = 2; rows <= 60; rows += 1) {
        const out = layoutChart(chartOf("bar", rows, seriesCount), BOX);
        // A row's share per bar, read back off the bars as drawn.
        if (Math.abs(rowShare(out) - LABEL_HEIGHT) < EDGE) continue;
        expect(
          valueLabels(out).length > 0,
          `${String(rows)} rows x ${String(seriesCount)}`,
        ).toBe(rowShare(out) >= LABEL_HEIGHT);
      }
    }
  });
});

describe("a tornado", () => {
  const tornado = (rows: number) =>
    chartOf("bar", rows, 2, {
      overlap: true,
      series: chartOf("bar", rows, 2).series.map((one, j) => ({
        ...one,
        values: one.values.map((value) => (j === 0 ? -(value ?? 0) : value)),
        labels: one.labels.map((text) => (j === 0 ? `(${text})` : text)),
      })),
    });

  it("keeps its labels with a dozen drivers and drops them at forty", () => {
    expect(valueLabels(layoutChart(tornado(12), BOX)).length).toBeGreaterThan(
      0,
    );
    expect(valueLabels(layoutChart(tornado(40), BOX))).toHaveLength(0);
  });
});

describe("a stacked bar chart with too little room for its segment labels", () => {
  it("draws every segment and no label at 40 rows of 6 series", () => {
    const out = layoutChart(chartOf("stackedBar", 40, 6), BOX);
    expect(bars(out)).toHaveLength(240);
    expect(valueLabels(out)).toHaveLength(0);
  });

  it("draws every segment and no label at 120 rows of 3 series", () => {
    const out = layoutChart(chartOf("stackedBar", 120, 3), BOX);
    expect(bars(out)).toHaveLength(360);
    expect(valueLabels(out)).toHaveLength(0);
  });

  it("keeps its labels when the rows have room", () => {
    const out = layoutChart(chartOf("stackedBar", 5, 3), BOX);
    expect(valueLabels(out).length).toBeGreaterThan(0);
  });

  it("draws labels exactly while a row is a label high", () => {
    for (let rows = 2; rows <= 60; rows += 1) {
      const out = layoutChart(chartOf("stackedBar", rows, 3), BOX);
      if (Math.abs(rowShare(out) - LABEL_HEIGHT) < EDGE) continue;
      expect(valueLabels(out).length > 0, `${String(rows)} rows`).toBe(
        rowShare(out) >= LABEL_HEIGHT,
      );
    }
  });
});
