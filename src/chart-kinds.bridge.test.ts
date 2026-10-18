// A waterfall of fewer than three points has no bridge to draw: the payload
// validator still admits it (two points is the floor of every kind), so the kind
// rule declines it in words and its layout gives no primitives, where
// bridgeSeries used to throw out of every caller. Every other kind with two
// points, and a waterfall with three, are unchanged.

import { describe, expect, it } from "vitest";
import { chartDecline, kindDecline, layoutKind } from "./chart-kinds";
import { layoutChart, layoutClassic } from "./chart-shapes";
import { MIN_BRIDGE_POINTS } from "./chartmath";
import { CHART_KINDS, isChartData, type ChartKind } from "./link/chart-model";
import { chartOf } from "../test/hunt/ppt.chart-of.hunt3.support";

const BOX = { left: 0, top: 0, width: 480, height: 270 };

describe("a waterfall too short to bridge", () => {
  it("declines in words, for one point and for two", () => {
    expect(chartDecline(chartOf("waterfall", 2))).toBe(
      "2 points; waterfall shapes need at least 3",
    );
    expect(chartDecline(chartOf("waterfall", 1))).toBe(
      "1 point; waterfall shapes need at least 3",
    );
    expect(MIN_BRIDGE_POINTS).toBe(3);
  });

  it("is still chart data the payload validator accepts: the kind rules decide, not the decoder", () => {
    expect(isChartData(chartOf("waterfall", 2))).toBe(true);
    expect(kindDecline("waterfall")).toBeNull();
  });

  it("lays out no primitives, where it used to throw", () => {
    const two = chartOf("waterfall", 2);
    expect(layoutKind(two, BOX)).toBeNull();
    expect(layoutChart(two, BOX)).toEqual([]);
  });

  it("is drawn from three points exactly as the classic layout draws it", () => {
    const three = chartOf("waterfall", 3);
    expect(chartDecline(three)).toBeNull();
    expect(layoutChart(three, BOX)).toEqual(layoutClassic(three, BOX));
    expect(layoutChart(three, BOX).length).toBeGreaterThan(0);
  });

  it("leaves every other kind with two points to its own rules", () => {
    for (const kind of CHART_KINDS) {
      if (kind === "waterfall" || kind === "combo" || kind === "scatter") {
        continue;
      }
      const data = chartOf(kind as ChartKind, 2, kind === "pie" ? 1 : 2);
      expect(chartDecline(data), kind).toBeNull();
    }
  });
});
