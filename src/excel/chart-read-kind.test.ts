// The pure grammar of reading a chart: which Excel chart type is which kind (and
// whether it is a 100 % stack), what the series of one chart add up to as a kind
// (a column and a line, or a series on the secondary axis, make a combo), which
// flags a series carries, and when a line is the total of a stack.

import { describe, expect, it } from "vitest";
import {
  chartKind,
  classify,
  decideKind,
  foldTotals,
  seriesFlags,
  type SeriesFacts,
} from "./chart-read-kind";

function fact(chartType: string, extra: Partial<SeriesFacts> = {}) {
  return {
    chartType,
    axisGroup: "Primary",
    hasDataLabels: true,
    markerStyle: "Automatic",
    ...extra,
  } satisfies SeriesFacts;
}

describe("classify", () => {
  it("reads the column families, clustered, stacked and 100 % stacked", () => {
    expect(classify("ColumnClustered")).toEqual({
      kind: "column",
      percent: false,
    });
    expect(classify("3DColumnClustered")?.kind).toBe("column");
    expect(classify("CylinderColClustered")?.kind).toBe("column");
    expect(classify("ColumnStacked")).toEqual({
      kind: "stackedColumn",
      percent: false,
    });
    expect(classify("ColumnStacked100")).toEqual({
      kind: "stackedColumn",
      percent: true,
    });
    expect(classify("PyramidColStacked")?.kind).toBe("stackedColumn");
  });

  it("reads the bar families before the clustered catch-all", () => {
    expect(classify("BarClustered")?.kind).toBe("bar");
    expect(classify("ConeBarClustered")?.kind).toBe("bar");
    expect(classify("BarStacked")).toEqual({
      kind: "stackedBar",
      percent: false,
    });
    expect(classify("BarStacked100")).toEqual({
      kind: "stackedBar",
      percent: true,
    });
    expect(classify("3DBarStacked")?.kind).toBe("stackedBar");
  });

  it("reads area, stacked area and 100 % stacked area", () => {
    expect(classify("Area")).toEqual({ kind: "area", percent: false });
    expect(classify("3DArea")?.kind).toBe("area");
    expect(classify("AreaStacked")).toEqual({
      kind: "stackedArea",
      percent: false,
    });
    expect(classify("AreaStacked100")).toEqual({
      kind: "stackedArea",
      percent: true,
    });
    expect(classify("3DAreaStacked")?.kind).toBe("stackedArea");
  });

  it("reads the lines, the pies, the doughnuts, the scatters and the waterfall", () => {
    expect(classify("Line")?.kind).toBe("line");
    expect(classify("LineMarkers")?.kind).toBe("line");
    expect(classify("Pie")?.kind).toBe("pie");
    expect(classify("3DPieExploded")?.kind).toBe("pie");
    expect(classify("Doughnut")?.kind).toBe("doughnut");
    expect(classify("DoughnutExploded")?.kind).toBe("doughnut");
    expect(classify("XYScatter")?.kind).toBe("scatter");
    expect(classify("XYScatterLinesNoMarkers")?.kind).toBe("scatter");
    expect(classify("XYScatterSmooth")?.kind).toBe("scatter");
    expect(classify("Waterfall")?.kind).toBe("waterfall");
  });

  it("leaves every type with no shapes to the picture", () => {
    for (const type of [
      "",
      "Invalid",
      "Radar",
      "RadarFilled",
      "Surface",
      "Bubble",
      "StockHLC",
      "PieOfPie",
      "BarOfPie",
      "LineStacked",
      "LineMarkersStacked100",
      "Treemap",
      "Funnel",
    ]) {
      expect(classify(type)).toBeNull();
    }
  });
});

describe("chartKind", () => {
  it("answers the kind, and no kind for a type with none", () => {
    expect(chartKind("ColumnStacked100", null)).toBe("stackedColumn");
    expect(chartKind("Doughnut", null)).toBe("doughnut");
    expect(chartKind("Area", null)).toBe("area");
    expect(chartKind("XYScatter", null)).toBe("scatter");
    expect(chartKind("Radar", null)).toBeNull();
    expect(chartKind("", null)).toBeNull();
  });

  // Columns are laid out side by side; a full overlap would hide one behind
  // the other, so that chart is one the picture says better.
  it("refuses columns drawn on top of each other, and lets a tornado's bars be", () => {
    expect(chartKind("ColumnClustered", 100)).toBeNull();
    expect(chartKind("BarClustered", 100)).toBe("bar");
    expect(chartKind("ColumnClustered", 0)).toBe("column");
  });
});

describe("decideKind", () => {
  it("is the chart's own kind when every series is of one type", () => {
    expect(
      decideKind("ColumnClustered", [
        fact("ColumnClustered"),
        fact("ColumnClustered"),
      ]),
    ).toEqual({ kind: "column", percent: false, types: null });
    expect(decideKind("Line", [fact("Line")])).toEqual({
      kind: "line",
      percent: false,
      types: null,
    });
  });

  it("flags a 100 % stack from the type and keeps a plain one plain", () => {
    expect(
      decideKind("ColumnStacked100", [
        fact("ColumnStacked100"),
        fact("ColumnStacked100"),
      ]),
    ).toEqual({ kind: "stackedColumn", percent: true, types: null });
    expect(
      decideKind("ColumnStacked", [
        fact("ColumnStacked"),
        fact("ColumnStacked"),
      ]),
    ).toEqual({ kind: "stackedColumn", percent: false, types: null });
  });

  it("trusts the series when the chart's own type is only the first group's", () => {
    expect(
      decideKind("ColumnClustered", [fact("Line"), fact("LineMarkers")]),
    ).toEqual({ kind: "line", percent: false, types: null });
  });

  it("is the chart's own kind when the series answered nothing", () => {
    expect(decideKind("BarClustered", [fact(""), fact("")])).toEqual({
      kind: "bar",
      percent: false,
      types: null,
    });
  });

  it("is a combo when the types differ, naming each series' type", () => {
    expect(
      decideKind("ColumnClustered", [fact("ColumnClustered"), fact("Line")]),
    ).toEqual({
      kind: "combo",
      percent: false,
      types: ["column", "line"],
      percents: [false, false],
    });
    expect(
      decideKind("ColumnStacked", [
        fact("ColumnStacked"),
        fact("ColumnStacked"),
        fact("LineMarkers"),
      ]),
    ).toEqual({
      kind: "combo",
      percent: false,
      types: ["stackedColumn", "stackedColumn", "line"],
      percents: [false, false, false],
    });
  });

  it("is a combo when a series sits on the secondary axis, even of one type", () => {
    expect(
      decideKind("Line", [
        fact("Line"),
        fact("Line", { axisGroup: "Secondary" }),
      ]),
    ).toEqual({
      kind: "combo",
      percent: false,
      types: ["line", "line"],
      percents: [false, false],
    });
  });

  it("is not a combo when every series sits on the secondary axis: that is one axis", () => {
    expect(
      decideKind("Line", [
        fact("Line", { axisGroup: "Secondary" }),
        fact("Line", { axisGroup: "Secondary" }),
      ]),
    ).toEqual({ kind: "line", percent: false, types: null });
  });

  it("names a series type the slide has no shapes for", () => {
    expect(
      decideKind("ColumnClustered", [fact("ColumnClustered"), fact("Radar")]),
    ).toEqual({ issue: "Radar series are not drawn as shapes" });
  });

  it("names a chart type the slide has no shapes for", () => {
    expect(decideKind("Radar", [fact("Radar")])).toEqual({
      issue: "Radar charts are not drawn as shapes",
    });
    expect(decideKind("", [])).toEqual({
      issue: "this chart type is not drawn as shapes",
    });
  });

  it("leaves a waterfall to its own type whatever the series say", () => {
    expect(decideKind("Waterfall", [fact("")])).toEqual({
      kind: "waterfall",
      percent: false,
      types: null,
    });
  });

  it("keeps a doughnut and a scatter as they are", () => {
    expect(decideKind("Doughnut", [fact("Doughnut")])).toMatchObject({
      kind: "doughnut",
    });
    expect(decideKind("XYScatter", [fact("XYScatter")])).toMatchObject({
      kind: "scatter",
    });
  });
});

describe("seriesFlags", () => {
  it("says nothing for a series with its labels and markers on the primary axis", () => {
    expect(seriesFlags(fact("Line"), "line")).toEqual({});
    expect(seriesFlags(fact("ColumnClustered"), "column")).toEqual({});
  });

  it("flags a series without data labels", () => {
    expect(seriesFlags(fact("Line", { hasDataLabels: false }), "line")).toEqual(
      { dataLabels: false },
    );
    expect(
      seriesFlags(fact("ColumnClustered", { hasDataLabels: false }), "column"),
    ).toEqual({ dataLabels: false });
  });

  it("says nothing when the host did not say", () => {
    expect(
      seriesFlags(
        fact("Line", { hasDataLabels: null, markerStyle: "" }),
        "line",
      ),
    ).toEqual({});
  });

  it("flags a line without markers, and no other type", () => {
    expect(seriesFlags(fact("Line", { markerStyle: "None" }), "line")).toEqual({
      markers: false,
    });
    expect(
      seriesFlags(fact("Line", { markerStyle: "Circle" }), "line"),
    ).toEqual({});
    expect(
      seriesFlags(
        fact("XYScatterLinesNoMarkers", { markerStyle: "None" }),
        "scatter",
      ),
    ).toEqual({ markers: false });
    expect(
      seriesFlags(fact("ColumnClustered", { markerStyle: "None" }), "column"),
    ).toEqual({});
  });

  it("flags the secondary axis", () => {
    expect(
      seriesFlags(fact("Line", { axisGroup: "Secondary" }), "line"),
    ).toEqual({ axis: "secondary" });
  });
});

// A banker's total on top of a stack: a line series that sits on the stack's
// own sums, with its labels on. The chart is the stack, the line is its total.
describe("foldTotals", () => {
  const stack = (values: (number | null)[], name = "s") => ({
    name,
    type: "stackedColumn" as const,
    values,
    labels: values.map(String),
    dataLabels: true,
    secondary: false,
    percent: false,
  });
  const total = (values: (number | null)[], labels = values.map(String)) => ({
    name: "Total",
    type: "line" as const,
    values,
    labels,
    dataLabels: true,
    secondary: false,
    percent: false,
  });

  it("takes the line that sums the stack as its totals and drops it", () => {
    const folded = foldTotals([
      stack([1, 2, 3]),
      stack([4, 5, 6]),
      total([5, 7, 9], ["5.0", "7.0", "9.0"]),
    ]);
    expect(folded).toEqual({ keep: [0, 1], totals: ["5.0", "7.0", "9.0"] });
  });

  it("counts a gap in the stack as zero", () => {
    const folded = foldTotals([
      stack([1, null, 3]),
      stack([4, 5, 6]),
      total([5, 5, 9]),
    ]);
    expect(folded?.keep).toEqual([0, 1]);
  });

  it("finds the line wherever it sits in the series order", () => {
    const folded = foldTotals([
      total([5, 7, 9]),
      stack([1, 2, 3]),
      stack([4, 5, 6]),
    ]);
    expect(folded?.keep).toEqual([1, 2]);
  });

  it("is not a total when it does not sum the stack", () => {
    expect(
      foldTotals([stack([1, 2, 3]), stack([4, 5, 6]), total([5, 7, 10])]),
    ).toBeNull();
  });

  it("is not a total when the line has no labels to show", () => {
    expect(
      foldTotals([
        stack([1, 2, 3]),
        stack([4, 5, 6]),
        { ...total([5, 7, 9]), dataLabels: false },
      ]),
    ).toBeNull();
  });

  it("is not a total when the line is on the secondary axis", () => {
    expect(
      foldTotals([
        stack([1, 2, 3]),
        stack([4, 5, 6]),
        { ...total([5, 7, 9]), secondary: true },
      ]),
    ).toBeNull();
  });

  it("is not a total when there is more than one line, or none", () => {
    expect(
      foldTotals([stack([1, 2, 3]), total([1, 2, 3]), total([1, 2, 3])]),
    ).toBeNull();
    expect(foldTotals([stack([1, 2, 3]), stack([4, 5, 6])])).toBeNull();
  });

  it("is not a total when the rest is not a plain stack", () => {
    expect(
      foldTotals([
        { ...stack([1, 2, 3]), type: "column" as const },
        total([1, 2, 3]),
      ]),
    ).toBeNull();
  });

  it("gives a category the line has no value for no total label", () => {
    const folded = foldTotals([
      stack([1, 2, 3]),
      stack([4, 5, 6]),
      total([5, null, 9], ["5", "", "9"]),
    ]);
    expect(folded?.totals).toEqual(["5", "", "9"]);
  });
});
