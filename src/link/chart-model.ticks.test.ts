// The tick-position flag of a chart payload: valid on the kinds that draw
// points along a category axis (a line, an area, a stacked area) and nowhere
// else, and read by the frozen v2.9.25 validator, so an older pane keeps the
// chart and draws it between the ticks as it always did.

import { describe, expect, it } from "vitest";
import { isChartData as frozenIsChartData } from "../../test/frozen/chart-model.v2.9.25";
import { isChartData, type ChartData, type ChartKind } from "./chart-model";

function chart(kind: ChartKind, extra: object = {}): ChartData {
  return {
    v: 1,
    kind,
    title: null,
    categories: ["a", "b", "c"],
    series: [
      {
        name: "s",
        values: [1, 2, 3],
        labels: ["1", "2", "3"],
        colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
      },
    ],
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
    ...extra,
  };
}

describe("ChartData.onTicks", () => {
  it("is valid on a line, an area and a stacked area, and a payload without it is too", () => {
    for (const kind of ["line", "area", "stackedArea"] as const) {
      expect(isChartData(chart(kind, { onTicks: true }))).toBe(true);
      expect(isChartData(chart(kind))).toBe(true);
    }
  });

  it("is refused on any other kind, and when it is not true", () => {
    for (const kind of ["column", "stackedBar", "bar", "waterfall", "pie"]) {
      expect(isChartData(chart(kind as ChartKind, { onTicks: true }))).toBe(
        false,
      );
    }
    expect(isChartData(chart("line", { onTicks: false }))).toBe(false);
    expect(isChartData(chart("area", { onTicks: "yes" }))).toBe(false);
  });

  it("is read by the frozen v2.9.25 validator on a line, flag and all", () => {
    expect(frozenIsChartData(chart("line", { onTicks: true }))).toBe(true);
  });
});
