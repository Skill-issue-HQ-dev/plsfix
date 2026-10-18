// The series half of the chart payload: what one series may carry, the
// defaults an absent field stands for (so a payload from an older pane draws as
// it always did), and the one place that knows every per-point array of a chart
// (so a reorder can never leave one behind).

import { describe, expect, it } from "vitest";
import type { ChartData } from "./chart-model";
import {
  filled,
  isSecondary,
  isSeries,
  positiveTotal,
  reindexPoints,
  seriesType,
  showsLabels,
  showsMarkers,
  type ChartSeries,
} from "./chart-series";

const one: ChartSeries = {
  name: "Revenue",
  values: [1, 2, 3],
  labels: ["1", "2", "3"],
  colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
};

const chart: ChartData = {
  v: 1,
  kind: "stackedColumn",
  title: null,
  categories: ["a", "b", "c"],
  series: [
    { ...one, x: [10, 20, null] },
    {
      name: "Cost",
      values: [4, null, 6],
      labels: ["4", "", "6"],
      colors: ["#111111", "#222222", "#333333"],
    },
  ],
  totals: ["5", "2", "9"],
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
};

describe("isSeries", () => {
  it("accepts a plain series and one with every optional field", () => {
    expect(isSeries(one, 3)).toBe(true);
    expect(
      isSeries(
        {
          ...one,
          type: "line",
          axis: "secondary",
          dataLabels: false,
          markers: false,
          x: [1, null, 3],
        },
        3,
      ),
    ).toBe(true);
  });

  it("holds every per-point array to the point count", () => {
    expect(isSeries({ ...one, values: [1, 2] }, 3)).toBe(false);
    expect(isSeries({ ...one, labels: ["1"] }, 3)).toBe(false);
    expect(isSeries({ ...one, colors: ["#2EC4B6"] }, 3)).toBe(false);
    expect(isSeries({ ...one, x: [1] }, 3)).toBe(false);
  });

  it("refuses a name that is not a string and a value that is not a number or null", () => {
    expect(isSeries({ ...one, name: 7 }, 3)).toBe(false);
    expect(isSeries({ ...one, values: [1, "2", 3] }, 3)).toBe(false);
    expect(isSeries(null, 3)).toBe(false);
  });
});

describe("what an absent field stands for", () => {
  it("takes the chart's own type when the series names none", () => {
    expect(seriesType(chart, one)).toBe("stackedColumn");
    expect(seriesType({ ...chart, kind: "line" }, one)).toBe("line");
    expect(seriesType({ ...chart, kind: "bar" }, one)).toBe("bar");
    expect(seriesType({ ...chart, kind: "stackedArea" }, one)).toBe(
      "stackedArea",
    );
  });

  it("takes the series' own type when it names one", () => {
    expect(
      seriesType({ ...chart, kind: "combo" }, { ...one, type: "line" }),
    ).toBe("line");
  });

  it("draws labels and markers unless the series says it has none", () => {
    expect(showsLabels(one)).toBe(true);
    expect(showsMarkers(one)).toBe(true);
    expect(showsLabels({ ...one, dataLabels: false })).toBe(false);
    expect(showsMarkers({ ...one, markers: false })).toBe(false);
    expect(showsLabels({ ...one, dataLabels: true })).toBe(true);
  });

  it("is on the primary axis unless it says secondary", () => {
    expect(isSecondary(one)).toBe(false);
    expect(isSecondary({ ...one, axis: "primary" })).toBe(false);
    expect(isSecondary({ ...one, axis: "secondary" })).toBe(true);
  });

  it("fills a gap with zero for the maths that cannot hold one", () => {
    expect(filled([1, null, 3])).toEqual([1, 0, 3]);
    expect(filled([])).toEqual([]);
  });

  it("divides a pie by its positive values only", () => {
    expect(positiveTotal([1, null, -5, 3, 0])).toBe(4);
    expect(positiveTotal([null, -1])).toBe(0);
    expect(positiveTotal([])).toBe(0);
  });
});

describe("a smooth series", () => {
  it("is a series that says so with a boolean, and only a boolean", () => {
    expect(isSeries({ ...one, smooth: true }, 3)).toBe(true);
    expect(isSeries({ ...one, smooth: false }, 3)).toBe(true);
    expect(isSeries({ ...one, smooth: "yes" }, 3)).toBe(false);
    expect(isSeries({ ...one, smooth: 1 }, 3)).toBe(false);
    expect(isSeries(one, 3)).toBe(true);
  });

  it("keeps its flag when the points are reordered", () => {
    const smooth: ChartData = {
      ...chart,
      series: [{ ...one, smooth: true }, chart.series[1]!],
    };
    const turned = reindexPoints(smooth, [2, 1, 0]);
    expect(turned.series[0]!.smooth).toBe(true);
    expect("smooth" in turned.series[1]!).toBe(false);
  });
});

describe("reindexPoints", () => {
  const reversed = reindexPoints(chart, [2, 1, 0]);

  it("reorders the categories and every array of every series", () => {
    expect(reversed.categories).toEqual(["c", "b", "a"]);
    expect(reversed.series[0]!.values).toEqual([3, 2, 1]);
    expect(reversed.series[0]!.labels).toEqual(["3", "2", "1"]);
    expect(reversed.series[0]!.x).toEqual([null, 20, 10]);
    expect(reversed.series[1]!.values).toEqual([6, null, 4]);
    expect(reversed.series[1]!.colors).toEqual([
      "#333333",
      "#222222",
      "#111111",
    ]);
  });

  it("reorders the totals with the categories", () => {
    expect(reversed.totals).toEqual(["9", "2", "5"]);
  });

  it("keeps everything that is not per point, and the original untouched", () => {
    expect(reversed).toMatchObject({
      kind: "stackedColumn",
      font: "Arial",
      ink: "#333333",
      titleColor: "#14213D",
    });
    expect(chart.categories).toEqual(["a", "b", "c"]);
    expect(chart.series[0]!.values).toEqual([1, 2, 3]);
  });

  it("is its own inverse for a reversal", () => {
    expect(reindexPoints(reversed, [2, 1, 0])).toEqual(chart);
  });

  it("leaves a series without x or a chart without totals without them", () => {
    const plain = reindexPoints({ ...chart, totals: undefined }, [2, 1, 0]);
    expect("totals" in plain).toBe(false);
    expect("x" in plain.series[1]!).toBe(false);
  });
});
