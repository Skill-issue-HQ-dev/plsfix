// A chart whose own type Excel for the web will not name. The first real run of
// the combo flow (05.10.2026) read a combo's Chart.chartType as null while its
// series answered ColumnClustered on the primary axis and LineMarkers on the
// secondary one: the chart must then be decided from its series, shipped as a
// combo with its data, and one that nothing can decide keeps the picture with
// a sentence that never says "null". Strict load semantics are on; the host is
// the fake's web-shaped one.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { pictureNote } from "../src/link/chart-model";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["2024A", "2025E", "2026E", "2027E"];

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost({ web: true });
});
afterEach(() => uninstallFakeHost());

const series = (name: string, base: number, extra: object = {}) => ({
  name,
  categories: NAMES,
  values: NAMES.map((_one, i) => base + i * 10),
  ...extra,
});

// What a user must never be shown for a chart type nobody named: the words null
// and undefined, or a type name that is empty (a sentence led or doubled by a gap).
function saysNoMissingType(text: string): void {
  expect(text).not.toMatch(/\bnull\b|\bundefined\b/i);
  expect(text).not.toMatch(/^\s|\s{2}/);
}

describe("exportActiveChart on a chart the web gives no type", () => {
  it("ships a column and a secondary-axis line as a combo with its chart data", async () => {
    const { payload, result } = await exportChart(boot, null, [
      series("Revenue", 100, { chartType: "ColumnClustered" }),
      series("Margin", 5, {
        chartType: "LineMarkers",
        axisGroup: "Secondary",
      }),
    ]);
    expect(payload.chartIssue).toBeUndefined();
    expect(payload.chart?.kind).toBe("combo");
    expect(payload.chart?.series.map((one) => one.type)).toEqual([
      "column",
      "line",
    ]);
    expect(payload.chart?.series.map((one) => one.axis)).toEqual([
      undefined,
      "secondary",
    ]);
    expect(payload.chart?.series[0]?.values).toEqual([100, 110, 120, 130]);
    expect(result.note).toBeUndefined();
  });

  it("reads the stacked columns of a combo the same way", async () => {
    const { payload } = await exportChart(boot, null, [
      series("Retail", 60, { chartType: "ColumnStacked" }),
      series("Corporate", 40, { chartType: "ColumnStacked" }),
      series("Margin", 5, { chartType: "LineMarkers", axisGroup: "Secondary" }),
    ]);
    expect(payload.chart?.kind).toBe("combo");
    expect(payload.chart?.series.map((one) => one.type)).toEqual([
      "stackedColumn",
      "stackedColumn",
      "line",
    ]);
  });

  it("keeps a combo of what has no shapes a picture, with the reason in words", async () => {
    const { payload, result } = await exportChart(boot, null, [
      series("Revenue", 100, { chartType: "ColumnClustered" }),
      series("Cash", 5, { chartType: "Area" }),
    ]);
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(
      "area series in a combo chart are not drawn as shapes",
    );
    expect(result.note).toBe(pictureNote(payload.chartIssue!));
  });

  it("still sees columns drawn on top of each other behind a null type", async () => {
    const { payload } = await exportChart(boot, null, [
      series("Revenue", 100, { chartType: "ColumnClustered", overlap: 100 }),
      series("Margin", 5, { chartType: "LineMarkers", axisGroup: "Secondary" }),
    ]);
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(
      "overlapping columns are not drawn as shapes",
    );
  });

  it("is a picture, in words that never say null, when the series cannot decide either", async () => {
    const { payload, result } = await exportChart(boot, null, [
      series("Revenue", 100, { chartType: "" }),
      series("Margin", 5, { chartType: "", axisGroup: "Secondary" }),
    ]);
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe("this chart type is not drawn as shapes");
    expect(result.note).toBe(
      "as a picture: this chart type is not drawn as shapes",
    );
    saysNoMissingType(payload.chartIssue!);
    saysNoMissingType(result.note!);
  });

  it("names the one series type it cannot draw, not the missing chart type", async () => {
    const { payload } = await exportChart(boot, null, [
      series("Revenue", 100, { chartType: "ColumnClustered" }),
      series("Mix", 5, { chartType: "RadarFilled" }),
    ]);
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(
      "RadarFilled series are not drawn as shapes",
    );
  });
});
