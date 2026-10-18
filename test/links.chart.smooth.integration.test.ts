// A smooth line chart exported from Excel against the strict Excel fake: the
// series Excel draws as a curve is flagged in the payload, the export row says
// the slide draws it straight (the same sentence the PowerPoint row uses), a
// straight line and every other kind says nothing, and the flag rides the
// batch the other series facts do: no round trip more, lost with them when the
// host refuses them.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { SMOOTH_STRAIGHT } from "../src/link/chart-model";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["Q1", "Q2", "Q3", "Q4"];

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost();
});
afterEach(() => uninstallFakeHost());

const series = (name: string, base: number, extra: object = {}) => ({
  name,
  categories: NAMES,
  values: NAMES.map((_one, i) => base + i * 10),
  ...extra,
});

describe("exportActiveChart: a smooth line", () => {
  it("flags the series Excel draws as a curve and says the slide draws it straight", async () => {
    const { payload, result } = await exportChart(boot, "LineMarkers", [
      series("A", 10, { smooth: true }),
      series("B", 20),
    ]);
    expect(payload.chart?.kind).toBe("line");
    expect(payload.chart?.series[0]?.smooth).toBe(true);
    expect(payload.chart?.series[1]).not.toHaveProperty("smooth");
    expect(result.note).toBe(SMOOTH_STRAIGHT);
    expect(result.note).toBe("smooth lines are drawn straight");
  });

  it("says nothing about a straight line, and writes no flag", async () => {
    const { payload, result } = await exportChart(boot, "Line", [
      series("A", 10, { smooth: false }),
      series("B", 20),
    ]);
    expect(payload.chart?.series[0]).not.toHaveProperty("smooth");
    expect(payload.chart?.series[1]).not.toHaveProperty("smooth");
    expect(result.note).toBeUndefined();
  });

  it("reads the flag of a line chart only: a column series is never flagged", async () => {
    const { payload, result } = await exportChart(boot, "ColumnClustered", [
      series("A", 10, { smooth: true }),
    ]);
    expect(payload.chart?.kind).toBe("column");
    expect(payload.chart?.series[0]).not.toHaveProperty("smooth");
    expect(result.note).toBeUndefined();
  });

  it("rides the batch of the other series facts: not one sync more than a straight line costs", async () => {
    const before = boot.helpers.syncCount();
    await exportChart(boot, "Line", [series("A", 10)]);
    const straight = boot.helpers.syncCount() - before;
    const mark = boot.helpers.syncCount();
    await exportChart(boot, "Line", [series("A", 10, { smooth: true })]);
    expect(boot.helpers.syncCount() - mark).toBe(straight);
  });

  it("is lost with the other facts when the host refuses them, and the chart still ships", async () => {
    boot.helpers.failNextAxisRead();
    const { payload, result } = await exportChart(boot, "Line", [
      series("A", 10, { smooth: true }),
    ]);
    expect(payload.chart?.kind).toBe("line");
    expect(payload.chart?.series[0]).not.toHaveProperty("smooth");
    expect(result.note).toBeUndefined();
  });
});

describe("exportActiveChart: a host that refuses the smooth flag", () => {
  const facts = {
    hasDataLabels: false,
    markerStyle: "None",
    axisGroup: "Secondary",
  };

  it("is asked again without the smooth flag alone: the labels flag, markers and secondary axis stay", async () => {
    boot.helpers.failNextSmoothRead();
    const { payload } = await exportChart(boot, "Line", [
      series("A", 10, { smooth: true, ...facts }),
      series("B", 20),
    ]);
    const first = payload.chart?.series[0];
    expect(first).not.toHaveProperty("smooth");
    expect(first?.dataLabels).toBe(false);
    expect(first?.markers).toBe(false);
    expect(first?.axis).toBe("secondary");
  });

  it("costs one sync more than a host that answers, no more", async () => {
    const mark = boot.helpers.syncCount();
    await exportChart(boot, "Line", [series("A", 10, { smooth: true })]);
    const answered = boot.helpers.syncCount() - mark;
    boot.helpers.failNextSmoothRead();
    const refusedMark = boot.helpers.syncCount();
    await exportChart(boot, "Line", [series("A", 10, { smooth: true })]);
    expect(boot.helpers.syncCount() - refusedMark).toBe(answered + 1);
  });

  it("falls back to the dimensions alone when the host refuses the retry as well", async () => {
    boot.helpers.failNextSmoothRead();
    // Both attempts that touch the category axis fail; the third asks for the
    // dimensions only.
    boot.helpers.failNextAxisRead();
    boot.helpers.failNextAxisRead();
    const { payload } = await exportChart(boot, "Line", [
      series("A", 10, { smooth: true, ...facts }),
      series("B", 20),
    ]);
    expect(payload.chart?.kind).toBe("line");
    const first = payload.chart?.series[0];
    expect(first).not.toHaveProperty("smooth");
    expect(first).not.toHaveProperty("dataLabels");
    expect(first).not.toHaveProperty("axis");
  });
});
