// The tick position of a line or an area exported from Excel against the strict
// fake: a category axis Excel has on its ticks is flagged in the payload, one
// between them (a line's default) is not, a column never reads it, and a host
// that refuses the read loses the flag alone: the chart still ships, with the
// other series facts, and no export fails.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["Q1", "Q2", "Q3", "Q4"];
const series = (extra: object = {}) => [
  {
    name: "A",
    categories: NAMES,
    values: NAMES.map((_one, i) => 10 + i * 10),
    ...extra,
  },
];
const axis = (isBetweenCategories: boolean) => ({
  axes: { category: { isBetweenCategories }, value: {} },
});

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost();
});
afterEach(() => uninstallFakeHost());

describe("exportActiveChart: the category axis on its tick marks", () => {
  it("flags an area, a stacked area and a line whose axis Excel puts on the ticks", async () => {
    for (const [type, kind] of [
      ["Area", "area"],
      ["AreaStacked", "stackedArea"],
      ["Line", "line"],
    ] as const) {
      const { payload } = await exportChart(boot, type, series(), axis(false));
      expect(payload.chart?.kind).toBe(kind);
      expect(payload.chart?.onTicks).toBe(true);
    }
  });

  it("writes nothing for an axis between the ticks, and never reads it for a column", async () => {
    const between = await exportChart(boot, "Area", series(), axis(true));
    expect(between.payload.chart?.kind).toBe("area");
    expect(between.payload.chart).not.toHaveProperty("onTicks");
    const column = await exportChart(
      boot,
      "ColumnClustered",
      series(),
      axis(false),
    );
    expect(column.payload.chart?.kind).toBe("column");
    expect(column.payload.chart).not.toHaveProperty("onTicks");
  });

  it("rides the batch of the other facts: not one sync more than an axis between the ticks", async () => {
    const mark = boot.helpers.syncCount();
    await exportChart(boot, "Area", series(), axis(true));
    const between = boot.helpers.syncCount() - mark;
    const again = boot.helpers.syncCount();
    await exportChart(boot, "Area", series(), axis(false));
    expect(boot.helpers.syncCount() - again).toBe(between);
  });

  it("is lost alone when the host refuses it: the chart still ships with its other facts, one sync more", async () => {
    const mark = boot.helpers.syncCount();
    await exportChart(boot, "Area", series(), axis(false));
    const answered = boot.helpers.syncCount() - mark;
    boot.helpers.failNextBetweenRead();
    const refused = boot.helpers.syncCount();
    const { payload, result } = await exportChart(
      boot,
      "Area",
      series({ hasDataLabels: false }),
      axis(false),
    );
    expect(payload.chart?.kind).toBe("area");
    expect(payload.chart).not.toHaveProperty("onTicks");
    expect(payload.chart?.series[0]?.dataLabels).toBe(false);
    expect(result.note).toBeUndefined();
    expect(boot.helpers.syncCount() - refused).toBe(answered + 1);
  });
});
