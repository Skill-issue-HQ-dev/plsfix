// The chart data a link can carry: what passes the validator, what the caps
// refuse, how a blank cell travels (as null, a gap, never 0), the optional
// fields a newer pane adds (percent, categoriesReversed, totals, per-series
// type, axis and flags) and how an angle is brought onto the range PowerPoint
// reads back. Payloads from the shipped codec are in chart-model.shipped.test.ts.

import { describe, expect, it } from "vitest";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  chartCapIssue,
  INSERT_CANCELLED,
  INSERT_STOPPED,
  isChartData,
  needsNewerPowerPoint,
  NO_SHAPE_API,
  normalizeAngle,
  PICTURE_CHOSEN,
  pictureNote,
  REFRESH_STOPPED,
  type ChartData,
  type ChartKind,
} from "./chart-model";

const column: ChartData = {
  v: 1,
  kind: "column",
  title: "Revenue",
  categories: ["2024A", "2025E"],
  series: [
    {
      name: "Revenue",
      values: [100, 120],
      labels: ["100", "120"],
      colors: ["#2EC4B6", "#2EC4B6"],
    },
  ],
  font: "Arial",
  ink: "#333333",
  titleColor: "#14213D",
};

const series = column.series[0]!;

// A chart of `points` categories and `count` series of any kind: what every
// cap is measured on.
function chartOf(
  kind: ChartKind,
  points: number,
  count = 1,
  extra: Partial<ChartData> = {},
): ChartData {
  const categories = Array.from({ length: points }, (_, i) => `C${String(i)}`);
  return {
    ...column,
    kind,
    categories,
    series: Array.from({ length: count }, (_, j) => ({
      name: `S${String(j)}`,
      values: categories.map((_c, i) => 10 + i),
      labels: categories.map((_c, i) => String(10 + i)),
      colors: categories.map(() => "#2EC4B6"),
    })),
    ...extra,
  };
}

describe("isChartData", () => {
  it("accepts a consistent column chart", () => {
    expect(isChartData(column)).toBe(true);
  });

  it("refuses a series whose labels do not match its values", () => {
    expect(
      isChartData({ ...column, series: [{ ...series, labels: ["100"] }] }),
    ).toBe(false);
  });

  it("refuses fewer than two points and more than the cap", () => {
    expect(isChartData(chartOf("column", CHART_MAX_POINTS))).toBe(true);
    expect(isChartData(chartOf("column", CHART_MAX_POINTS + 1))).toBe(false);
    expect(isChartData(chartOf("column", 1))).toBe(false);
  });

  it("refuses a pie with two series or more slices than the cap", () => {
    expect(isChartData({ ...column, kind: "pie" })).toBe(true);
    expect(
      isChartData({ ...column, kind: "pie", series: [series, series] }),
    ).toBe(false);
    expect(isChartData(chartOf("pie", CHART_MAX_SLICES))).toBe(true);
    expect(isChartData(chartOf("pie", CHART_MAX_SLICES + 1))).toBe(false);
  });

  it("refuses an unknown kind, a bad colour, a non-finite value and a long title", () => {
    expect(isChartData({ ...column, kind: "line" })).toBe(true);
    expect(isChartData({ ...column, kind: "radar" })).toBe(false);
    expect(
      isChartData({
        ...column,
        series: [{ ...series, colors: ["red", "#2EC4B6"] }],
      }),
    ).toBe(false);
    expect(
      isChartData({
        ...column,
        series: [{ ...series, values: [100, Number.NaN] }],
      }),
    ).toBe(false);
    expect(
      isChartData({
        ...column,
        series: [{ ...series, values: [100, Number.POSITIVE_INFINITY] }],
      }),
    ).toBe(false);
    expect(isChartData({ ...column, title: "x".repeat(81) })).toBe(false);
    expect(isChartData({ ...column, title: null })).toBe(true);
  });

  it("allows overlap only as true", () => {
    expect(isChartData({ ...column, kind: "bar", overlap: true })).toBe(true);
    expect(isChartData({ ...column, kind: "bar", overlap: false })).toBe(false);
  });

  it("refuses anything that is not a record", () => {
    expect(isChartData(null)).toBe(false);
    expect(isChartData("chart")).toBe(false);
  });
});

// A blank cell is a gap in the data, never a zero: it travels as null.
describe("gaps", () => {
  it("accepts a null value and keeps its label and colour slot", () => {
    const gap: ChartData = {
      ...column,
      series: [{ ...series, values: [100, null], labels: ["100", ""] }],
    };
    expect(isChartData(gap)).toBe(true);
  });

  it("accepts a series that is nothing but gaps", () => {
    expect(
      isChartData({ ...column, series: [{ ...series, values: [null, null] }] }),
    ).toBe(true);
  });

  it("still refuses a string or undefined where a number or null belongs", () => {
    const bad = (values: unknown[]) =>
      isChartData({ ...column, series: [{ ...series, values }] });
    expect(bad([100, "120"])).toBe(false);
    expect(bad([100, undefined])).toBe(false);
    expect(bad([100])).toBe(false);
  });
});

describe("caps", () => {
  it("is 120 points, 12 series and 24 slices", () => {
    expect(CHART_MAX_POINTS).toBe(120);
    expect(CHART_MAX_SERIES).toBe(12);
    expect(CHART_MAX_SLICES).toBe(24);
  });

  it("accepts twelve series and refuses a thirteenth", () => {
    expect(isChartData(chartOf("column", 3, CHART_MAX_SERIES))).toBe(true);
    expect(isChartData(chartOf("column", 3, CHART_MAX_SERIES + 1))).toBe(false);
  });

  it("holds a full-size chart: 120 points and 12 series", () => {
    expect(
      isChartData(chartOf("line", CHART_MAX_POINTS, CHART_MAX_SERIES)),
    ).toBe(true);
  });
});

describe("the kinds a newer pane adds", () => {
  it("accepts a doughnut like a pie: up to 24 slices, and a ring for each series (the rings are in chart-model.rings.test.ts)", () => {
    expect(isChartData(chartOf("doughnut", 8))).toBe(true);
    expect(isChartData(chartOf("doughnut", CHART_MAX_SLICES))).toBe(true);
    expect(isChartData(chartOf("doughnut", CHART_MAX_SLICES + 1))).toBe(false);
    expect(isChartData(chartOf("doughnut", 4, 2))).toBe(true);
  });

  it("accepts a doughnut hole of 10 to 90 percent on a doughnut only", () => {
    expect(isChartData(chartOf("doughnut", 4, 1, { hole: 50 }))).toBe(true);
    expect(isChartData(chartOf("doughnut", 4, 1, { hole: 9 }))).toBe(false);
    expect(isChartData(chartOf("doughnut", 4, 1, { hole: 91 }))).toBe(false);
    expect(isChartData(chartOf("doughnut", 4, 1, { hole: 50.5 }))).toBe(false);
    expect(isChartData(chartOf("pie", 4, 1, { hole: 50 }))).toBe(false);
  });

  it("accepts area and stacked area like any category chart", () => {
    expect(isChartData(chartOf("area", 12, 3))).toBe(true);
    expect(isChartData(chartOf("stackedArea", 12, 3))).toBe(true);
  });

  it("needs an x per point on every scatter series", () => {
    const scatter = chartOf("scatter", 3);
    expect(isChartData(scatter)).toBe(false);
    const withX = {
      ...scatter,
      series: [{ ...scatter.series[0]!, x: [1, 2.5, null] }],
    };
    expect(isChartData(withX)).toBe(true);
    expect(
      isChartData({
        ...scatter,
        series: [{ ...scatter.series[0]!, x: [1, 2] }],
      }),
    ).toBe(false);
  });

  it("needs a type on every series of a combo, and at least two series", () => {
    const base = chartOf("combo", 3, 2);
    expect(isChartData(base)).toBe(false);
    const typed = {
      ...base,
      series: [
        { ...base.series[0]!, type: "column" as const },
        {
          ...base.series[1]!,
          type: "line" as const,
          axis: "secondary" as const,
        },
      ],
    };
    expect(isChartData(typed)).toBe(true);
    expect(isChartData({ ...typed, series: [typed.series[0]!] })).toBe(false);
  });

  it("refuses a series type or axis it does not know", () => {
    const bad = (extra: Record<string, unknown>) =>
      isChartData({ ...column, series: [{ ...series, ...extra }] });
    expect(bad({ type: "line" })).toBe(true);
    expect(bad({ type: "radar" })).toBe(false);
    expect(bad({ axis: "secondary" })).toBe(true);
    expect(bad({ axis: "tertiary" })).toBe(false);
  });
});

describe("percent, totals and the order of the categories", () => {
  it("allows percent only as true, and only on the stacked kinds", () => {
    expect(isChartData(chartOf("stackedColumn", 3, 2, { percent: true }))).toBe(
      true,
    );
    expect(isChartData(chartOf("stackedBar", 3, 2, { percent: true }))).toBe(
      true,
    );
    expect(isChartData(chartOf("stackedArea", 3, 2, { percent: true }))).toBe(
      true,
    );
    expect(isChartData(chartOf("column", 3, 2, { percent: true }))).toBe(false);
    expect(isChartData(chartOf("line", 3, 2, { percent: true }))).toBe(false);
    const wrong = { ...chartOf("stackedColumn", 3, 2), percent: false };
    expect(isChartData(wrong)).toBe(false);
  });

  it("allows categoriesReversed as a boolean on any kind", () => {
    expect(isChartData({ ...column, categoriesReversed: true })).toBe(true);
    expect(isChartData({ ...column, categoriesReversed: false })).toBe(true);
    expect(isChartData({ ...column, categoriesReversed: "yes" })).toBe(false);
  });

  it("allows one total label per category on a stacked column or bar", () => {
    const totals = ["100", "120", "90"];
    expect(isChartData(chartOf("stackedColumn", 3, 2, { totals }))).toBe(true);
    expect(isChartData(chartOf("stackedBar", 3, 2, { totals }))).toBe(true);
    expect(isChartData(chartOf("stackedColumn", 3, 2, { totals: ["1"] }))).toBe(
      false,
    );
    expect(isChartData(chartOf("column", 3, 2, { totals }))).toBe(false);
  });

  it("allows per-series label and marker flags as booleans", () => {
    const flagged = (extra: Record<string, unknown>) =>
      isChartData({ ...column, series: [{ ...series, ...extra }] });
    expect(flagged({ dataLabels: false, markers: false })).toBe(true);
    expect(flagged({ dataLabels: true, markers: true })).toBe(true);
    expect(flagged({ dataLabels: "no" })).toBe(false);
    expect(flagged({ markers: 0 })).toBe(false);
  });
});

describe("normalizeAngle", () => {
  it("maps onto (-180, 180] the way PowerPoint reads angles back", () => {
    expect(normalizeAngle(200)).toBe(-160);
    expect(normalizeAngle(269.9)).toBeCloseTo(-90.1);
    expect(normalizeAngle(360)).toBe(0);
    expect(normalizeAngle(-90)).toBe(-90);
    expect(normalizeAngle(180)).toBe(180);
    expect(normalizeAngle(-180)).toBe(180);
    expect(normalizeAngle(0)).toBe(0);
  });
});

// The sentence the panes show after "as a picture": the first cap the chart
// is outside of, counted the way the modeller sees the chart.
describe("chartCapIssue", () => {
  it("is null for a chart inside every cap", () => {
    expect(chartCapIssue("column", 2, 1)).toBeNull();
    expect(chartCapIssue("pie", 24, 1)).toBeNull();
    expect(chartCapIssue("doughnut", 24, 1)).toBeNull();
    expect(chartCapIssue("line", 120, 12)).toBeNull();
  });

  it("names the series, point and slice caps", () => {
    expect(chartCapIssue("column", 2, 0)).toBe("no series");
    expect(chartCapIssue("column", 2, 13)).toBe(
      "13 series; shapes draw up to 12",
    );
    expect(chartCapIssue("column", 1, 1)).toBe(
      "1 point; shapes need at least 2",
    );
    expect(chartCapIssue("column", 121, 1)).toBe(
      "121 points; shapes draw up to 120",
    );
    expect(chartCapIssue("pie", 3, 2)).toBe(
      "a pie with 2 series; pie shapes draw one",
    );
    expect(chartCapIssue("pie", 25, 1)).toBe(
      "25 slices; pie shapes draw up to 24",
    );
    expect(chartCapIssue("doughnut", 25, 1)).toBe(
      "25 slices; doughnut shapes draw up to 24",
    );
  });
});

// The only pictures left on a slide have a reason the row can say, in these
// words: a host with no shape API, a kind that needs a newer one, a kind with
// no shape version yet (chart-kinds.ts), or the user's own choice. Budget,
// size, caps, blanks and timeouts are not on the list: a stalled host says what
// happened instead.
describe("the sentences of the honest declines", () => {
  it("says what a host below PowerPointApi 1.8 lacks", () => {
    expect(NO_SHAPE_API).toBe(
      "as a picture: this PowerPoint has no shape API (Microsoft 365 2504 / Mac 16.96 or newer)",
    );
  });

  it("names the kind that needs the newer shape API", () => {
    expect(needsNewerPowerPoint("pie")).toBe(
      "as a picture: pie shapes need Microsoft 365 2601 / Mac 16.105 or newer",
    );
    expect(needsNewerPowerPoint("area")).toBe(
      "as a picture: area shapes need Microsoft 365 2601 / Mac 16.105 or newer",
    );
  });

  it("every decline opens with the one 'as a picture' words", () => {
    for (const note of [NO_SHAPE_API, needsNewerPowerPoint("pie")]) {
      expect(note.startsWith(pictureNote(""))).toBe(true);
    }
  });

  it("the user's own picture and the stops are plain sentences", () => {
    expect(PICTURE_CHOSEN).toBe("Picture (your choice)");
    expect(INSERT_STOPPED).toBe(
      "PowerPoint stopped answering; nothing was inserted, the export is still in the Inbox",
    );
    expect(REFRESH_STOPPED).toBe(
      "PowerPoint stopped answering; the old chart stays",
    );
    expect(INSERT_CANCELLED).toBe(
      "Cancelled: nothing was inserted, the export is still in the Inbox",
    );
  });
});
