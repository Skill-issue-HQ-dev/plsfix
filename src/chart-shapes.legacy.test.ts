// Payloads v2.9.25 wrote (no category-order field, caps 40 / 6 / 12) laid out by
// the current layout: every primitive equal to what that version's layoutChart
// drew, so no deck already on a slide changes when it is updated. The golden is
// test/golden/chart-layout.v2.9.25.json, produced by running the v2.9.25 sources
// (61f6a19: src/chart-shapes.ts and what it imports) on the same charts.
// Two cases are re-pinned on purpose (their note in the golden says so): a
// column and a waterfall chart now keep the room of a label under the top of the
// plot, so the tallest bar's value label stays out of the title band
// (chart-shapes.headroom.test.ts); every other case is still v2.9.25's drawing.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  SHIPPED_STACKED,
  SHIPPED_TORNADO,
} from "../test/golden/shipped-payloads";
import { layoutChart, type Primitive } from "./chart-shapes";
import type { Box } from "./layout";
import { isChartData, type ChartData } from "./link/chart-model";

interface GoldenCase {
  name: string;
  note: string;
  chart: unknown;
  primitives: Primitive[];
}

interface Golden {
  box: Box;
  cases: GoldenCase[];
}

const golden: Golden = JSON.parse(
  readFileSync(
    join(process.cwd(), "test/golden/chart-layout.v2.9.25.json"),
    "utf8",
  ),
);

function chartOf(name: string): ChartData {
  const found = golden.cases.find((one) => one.name === name);
  if (found === undefined || !isChartData(found.chart)) {
    throw new Error(`golden case ${name} is missing or not a chart`);
  }
  return found.chart;
}

const BAR_CASES = ["tornadoShipped", "tornado", "clusteredBar", "stackedBar"];

describe("charts v2.9.25 wrote, laid out by the current layout", () => {
  it.each(golden.cases.map((one) => [one.name, one] as const))(
    "%s is drawn exactly as its golden pins it",
    (_name, one) => {
      expect(isChartData(one.chart)).toBe(true);
      expect(layoutChart(one.chart as ChartData, golden.box)).toEqual(
        one.primitives,
      );
    },
  );

  it("has a golden for every drawn kind v2.9.25 wrote", () => {
    const kinds = new Set(
      golden.cases.map((one) => (one.chart as ChartData).kind),
    );
    for (const kind of [
      "column",
      "stackedColumn",
      "bar",
      "stackedBar",
      "waterfall",
      "pie",
      "line",
    ]) {
      expect(kinds.has(kind as ChartData["kind"]), kind).toBe(true);
    }
  });

  it("is made from the very payloads that shipped", () => {
    const tornado = JSON.parse(SHIPPED_TORNADO).chart;
    const stacked = JSON.parse(SHIPPED_STACKED).chart;
    expect(chartOf("tornadoShipped")).toEqual(tornado);
    expect(chartOf("stackedColumn")).toEqual(stacked);
  });

  it("would catch a bar chart that flips: an explicit false is a different drawing", () => {
    for (const name of BAR_CASES) {
      const asWritten = layoutChart(chartOf(name), golden.box);
      const flipped = layoutChart(
        { ...chartOf(name), categoriesReversed: false },
        golden.box,
      );
      expect(flipped, name).not.toEqual(asWritten);
    }
  });
});
