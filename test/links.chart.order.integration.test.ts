// The category order of a chart link, exported against the strict Excel fake: a
// bar chart and a stacked bar chart carry the sheet's own answer (an explicit
// false for a default axis, true for a reversed one) so a pane can draw Excel's
// order; any other kind writes the flag only when it is true; an axis the host
// would not read leaves the flag out and the chart still ships. The refusal is
// also the strict-fake proof of the reader asking again without the extras.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["A", "B", "C", "D"];
const REVERSED = { category: { reversePlotOrder: true }, value: {} };

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost();
});
afterEach(() => uninstallFakeHost());

// Two series over four categories, the second series' extras as given.
function exported(chartType: string, extra: object = {}, seriesExtra = {}) {
  return exportChart(
    boot,
    chartType,
    [
      { name: "2025", categories: NAMES, values: [4, 3, 2, 1], ...seriesExtra },
      { name: "2026", categories: NAMES, values: [5, 4, 3, 2], ...seriesExtra },
    ],
    extra,
  );
}

describe("a bar chart carries the sheet's own category order", () => {
  it("writes an explicit false for a default axis", async () => {
    const { payload } = await exported("BarClustered");
    expect(payload.chart?.kind).toBe("bar");
    expect(payload.chart?.categoriesReversed).toBe(false);
  });

  it("writes an explicit false for a stacked bar chart too", async () => {
    const { payload } = await exported("BarStacked");
    expect(payload.chart?.kind).toBe("stackedBar");
    expect(payload.chart?.categoriesReversed).toBe(false);
  });

  it("writes true for a reversed axis, on a bar and a stacked bar", async () => {
    const bar = await exported("BarClustered", { axes: REVERSED });
    const stacked = await exported("BarStacked", { axes: REVERSED });
    expect(bar.payload.chart?.categoriesReversed).toBe(true);
    expect(stacked.payload.chart?.categoriesReversed).toBe(true);
  });

  it("writes nothing for a column chart on a default axis", async () => {
    const { payload } = await exported("ColumnClustered");
    expect(payload.chart?.kind).toBe("column");
    expect(payload.chart).not.toHaveProperty("categoriesReversed");
  });

  it("still writes true for a column chart whose axis is reversed", async () => {
    const { payload } = await exported("ColumnClustered", { axes: REVERSED });
    expect(payload.chart?.categoriesReversed).toBe(true);
  });
});

describe("a bar chart whose axis could not be read", () => {
  it("ships without the flag when the host lacks ExcelApi 1.7", async () => {
    boot.helpers.setSupported(
      (set, version) => !(set === "ExcelApi" && version === "1.7"),
    );
    const { payload } = await exported("BarClustered");
    expect(payload.chart?.kind).toBe("bar");
    expect(payload.chart).not.toHaveProperty("categoriesReversed");
  });

  it("is asked again without the extras when the host refuses the axis: one sync more, the chart still ships", async () => {
    const before = boot.helpers.syncCount();
    const clean = await exported("BarClustered");
    const cleanSyncs = boot.helpers.syncCount() - before;
    expect(clean.payload.chart?.categoriesReversed).toBe(false);

    const mark = boot.helpers.syncCount();
    boot.helpers.failNextAxisRead();
    const refused = await exported("BarClustered");
    const refusedSyncs = boot.helpers.syncCount() - mark;

    expect(refused.payload.chart?.kind).toBe("bar");
    expect(refused.payload.chart?.series).toHaveLength(2);
    expect(refused.payload.chart).not.toHaveProperty("categoriesReversed");
    expect(refused.result.note).toBeUndefined();
    expect(refusedSyncs).toBe(cleanSyncs + 1);
  });

  it("loses the series flags with the extras, and the chart still draws", async () => {
    boot.helpers.failNextAxisRead();
    const { payload } = await exported(
      "ColumnClustered",
      {},
      { hasDataLabels: false },
    );
    expect(payload.chart?.kind).toBe("column");
    expect(payload.chart?.series[0]).not.toHaveProperty("dataLabels");
  });
});
