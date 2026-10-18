// The doughnut in the chart model (C5): a ring for each series up to 12, 24
// slices a ring (a pie stays one series), the slice kinds named in one place,
// and the hole as Excel's doughnutHoleSize gives it - a whole percent of the
// doughnut's own size - read as a fraction of its outer diameter, 75 % when the
// sheet did not say.

import { describe, expect, it } from "vitest";
import {
  CHART_KINDS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  chartCapIssue,
  DOUGHNUT_HOLE_DEFAULT,
  holeFraction,
  isChartData,
  isSliced,
  type ChartData,
  type ChartKind,
} from "./chart-model";

// A chart of `points` slices and `rings` series of any kind.
function chartOf(
  kind: ChartKind,
  points: number,
  rings = 1,
  extra: Partial<ChartData> = {},
): ChartData {
  return {
    v: 1,
    kind,
    title: null,
    categories: Array.from({ length: points }, (_, i) => `C${String(i)}`),
    series: Array.from({ length: rings }, (_, j) => ({
      name: `S${String(j)}`,
      values: Array.from({ length: points }, () => 1),
      labels: Array.from({ length: points }, () => "1"),
      colors: Array.from({ length: points }, () => "#2EC4B6"),
    })),
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
    ...extra,
  };
}

describe("a doughnut's rings", () => {
  it("accepts one ring or several, up to 24 slices a ring; a pie stays one series", () => {
    expect(isChartData(chartOf("doughnut", 8))).toBe(true);
    expect(isChartData(chartOf("doughnut", CHART_MAX_SLICES))).toBe(true);
    expect(isChartData(chartOf("doughnut", CHART_MAX_SLICES + 1))).toBe(false);
    // Each series is a ring, the first one innermost.
    expect(isChartData(chartOf("doughnut", 4, CHART_MAX_SERIES))).toBe(true);
    expect(isChartData(chartOf("doughnut", 4, CHART_MAX_SERIES + 1))).toBe(
      false,
    );
    expect(isChartData(chartOf("pie", 4, 2))).toBe(false);
  });

  it("names a ring's cap in the rows' words: 12 series, 24 slices", () => {
    expect(chartCapIssue("doughnut", 3, 2)).toBeNull();
    expect(chartCapIssue("doughnut", 24, 12)).toBeNull();
    expect(chartCapIssue("doughnut", 3, 13)).toBe(
      "13 series; shapes draw up to 12",
    );
    expect(chartCapIssue("doughnut", 25, 2)).toBe(
      "25 slices; doughnut shapes draw up to 24",
    );
    expect(chartCapIssue("pie", 3, 2)).toBe(
      "a pie with 2 series; pie shapes draw one",
    );
  });

  it("names the slices of a pie and a doughnut, and no other kind", () => {
    for (const kind of CHART_KINDS) {
      expect(isSliced(kind), kind).toBe(kind === "pie" || kind === "doughnut");
    }
  });
});

describe("a doughnut's hole", () => {
  it("is a fraction of the outer diameter: Excel's percent over 100, 75 % when the sheet did not say", () => {
    // Excel's doughnutHoleSize is a percentage of the chart's size (Microsoft's
    // VBA and Office.js docs; OOXML holeSize, 10 to 90): the hole's diameter
    // over the doughnut's own, which is also the inner radius over the outer.
    for (const hole of [10, 55, 90]) {
      expect(holeFraction(chartOf("doughnut", 4, 1, { hole }))).toBeCloseTo(
        hole / 100,
        12,
      );
    }
    expect(DOUGHNUT_HOLE_DEFAULT).toBe(75);
    expect(holeFraction(chartOf("doughnut", 4))).toBeCloseTo(0.75, 12);
  });
});
