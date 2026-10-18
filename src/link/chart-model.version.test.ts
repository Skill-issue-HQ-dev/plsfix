// The chart payload's version. A 100 % stack is written as v 2: a pane that
// shipped before 100 % stacks (frozen in test/frozen) reads v 1 only, refuses
// it, and keeps the picture - it never draws the stack as absolute segments.
// This pane reads v 1 and v 2; a chart of a version it has never heard of is
// ahead of it ("update pls,fix"), not broken.

import { describe, expect, it } from "vitest";
import {
  guardChart as frozenGuard,
  CHART_UNREADABLE as FROZEN_UNREADABLE,
} from "../../test/frozen/chart-guard.v2.9.25";
import { isChartData as frozenIsChartData } from "../../test/frozen/chart-model.v2.9.25";
import {
  CHART_NEEDS_UPDATE,
  CHART_UNREADABLE,
  guardChart,
} from "./chart-guard";
import { CHART_VERSION, isChartData, type ChartData } from "./chart-model";
import { decodePayload, encodePayload, type PicturePayload } from "./model";

const STACK: ChartData = {
  v: 1,
  kind: "stackedColumn",
  title: "Revenue by segment",
  categories: ["2024", "2025", "2026"],
  series: [
    {
      name: "Retail",
      values: [60, 70, 80],
      labels: ["60", "70", "80"],
      colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    },
    {
      name: "Corporate",
      values: [40, 50, 20],
      labels: ["40", "50", "20"],
      colors: ["#14213D", "#14213D", "#14213D"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};
const PERCENT: ChartData = { ...STACK, v: 2, percent: true };

function picture(chart: unknown): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: "iVBORw0KGgo=",
    src: {
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: "2026-10-03T08:00:00.000Z",
    hash: "0".repeat(64),
    chart,
  } as PicturePayload;
}

describe("the chart payload's version", () => {
  it("is read at v 1 and v 2, and at nothing else", () => {
    expect(CHART_VERSION).toBe(2);
    expect(isChartData({ ...STACK, v: 1 })).toBe(true);
    expect(isChartData({ ...STACK, v: 2 })).toBe(true);
    for (const v of [0, 3, 1.5, "1", "2", null, undefined]) {
      expect(isChartData({ ...STACK, v }), String(v)).toBe(false);
    }
  });

  it("travels in the encoded payload as written", () => {
    const decoded = decodePayload(encodePayload(picture(PERCENT)));
    if (decoded.kind !== "picture") throw new Error("not a picture");
    expect(decoded.chart?.v).toBe(2);
    expect(decoded.chart?.percent).toBe(true);
    expect(decoded.chartIssue).toBeUndefined();
  });
});

describe("a 100 % stack reaching a pane from before 100 % stacks", () => {
  it("is refused by the frozen v2.9.25 validator, accepted by this one", () => {
    expect(frozenIsChartData(PERCENT)).toBe(false);
    expect(isChartData(PERCENT)).toBe(true);
    // The same stack as v 1 is the very thing that pane would draw as an
    // absolute stack: the reason the version moves.
    expect(frozenIsChartData({ ...PERCENT, v: 1 })).toBe(true);
  });

  it("keeps the picture and says 'chart data unreadable' in the frozen pane", () => {
    const guarded = frozenGuard(picture(PERCENT)) as Record<string, unknown>;
    expect(guarded.kind).toBe("picture");
    expect("chart" in guarded).toBe(false);
    expect(guarded.chartIssue).toBe(FROZEN_UNREADABLE);
  });

  it("draws as a chart in this pane", () => {
    const payload = picture(PERCENT);
    expect(guardChart(payload)).toBe(payload);
  });
});

describe("a chart of a version this pane has never heard of", () => {
  it("is ahead of it: update pls,fix, not 'unreadable'", () => {
    const guarded = guardChart(picture({ ...PERCENT, v: 3 })) as Record<
      string,
      unknown
    >;
    expect(guarded.chartIssue).toBe(CHART_NEEDS_UPDATE);
    expect("chart" in guarded).toBe(false);
  });

  it("is ahead whatever else changed in it", () => {
    const guarded = guardChart(
      picture({ ...PERCENT, v: 4, kind: "bubble", series: "new shape" }),
    ) as Record<string, unknown>;
    expect(guarded.chartIssue).toBe(CHART_NEEDS_UPDATE);
  });

  it("leaves a broken chart of a known version 'unreadable'", () => {
    const guarded = guardChart(picture({ ...STACK, series: "nope" })) as Record<
      string,
      unknown
    >;
    expect(guarded.chartIssue).toBe(CHART_UNREADABLE);
    const old = guardChart(picture({ ...STACK, v: 0 })) as Record<
      string,
      unknown
    >;
    expect(old.chartIssue).toBe(CHART_UNREADABLE);
  });
});
