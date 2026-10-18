// The audit of 05.10.2026 (a web combo declined as "null charts are not drawn as
// shapes"): no sentence a chart declines with names a type nobody gave. Every
// kind's own sentence and every combo's, each with a different series type among
// its series, is whole words: no null, no undefined, no blank type.

import { describe, expect, it } from "vitest";
import { chartDecline } from "./chart-kinds";
import {
  CHART_KINDS,
  type ChartData,
  type ChartKind,
} from "./link/chart-model";
import type { SeriesType } from "./link/chart-series";

const SERIES_TYPES: SeriesType[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "line",
  "area",
  "stackedArea",
  "scatter",
  "pie",
  "doughnut",
];

function chartOf(kind: ChartKind, second?: SeriesType): ChartData {
  const series = (name: string, type?: SeriesType) => ({
    name,
    values: [1, 2, 3],
    labels: ["1", "2", "3"],
    colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    ...(type === undefined ? {} : { type }),
  });
  return {
    v: 1,
    kind,
    title: "T",
    categories: ["A", "B", "C"],
    series: [series("a", "column"), series("b", second)],
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
  };
}

describe("chartDecline never says a missing word", () => {
  it("leaves no null, undefined or blank type in the sentence of any kind or any combo", () => {
    const declines = [
      ...CHART_KINDS.map((kind) => chartDecline(chartOf(kind))),
      ...SERIES_TYPES.map((type) => chartDecline(chartOf("combo", type))),
    ].filter((sentence): sentence is string => sentence !== null);
    expect(declines.length).toBeGreaterThan(0);
    for (const sentence of declines) {
      expect(sentence).not.toMatch(/\bnull\b|\bundefined\b/i);
      expect(sentence).not.toMatch(/^\s|\s{2}/);
    }
  });
});
