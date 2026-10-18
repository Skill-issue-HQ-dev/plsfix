// Payloads written by the codec that shipped as v2.9.25 (61f6a19) decode with
// their chart intact and no new field added, and pass the validator that now
// allows gaps. Their strings live in test/golden/shipped-payloads.ts.

import { describe, expect, it } from "vitest";
import {
  SHIPPED_STACKED,
  SHIPPED_TORNADO,
} from "../../test/golden/shipped-payloads";
import { isChartData } from "./chart-model";
import { decodePayload, type PicturePayload } from "./model";

function decodeShipped(json: string): PicturePayload {
  const payload = decodePayload(new TextEncoder().encode(json));
  if (payload.kind !== "picture") throw new Error("not a picture");
  return payload;
}

describe("a payload from the shipped codec", () => {
  it("decodes with its chart intact and no new field added", () => {
    const decoded = decodeShipped(SHIPPED_STACKED);
    expect(decoded.chartIssue).toBeUndefined();
    expect(decoded.chart).toEqual(JSON.parse(SHIPPED_STACKED).chart);
    const chart = decoded.chart!;
    expect(chart.percent).toBeUndefined();
    expect(chart.totals).toBeUndefined();
    expect(chart.categoriesReversed).toBeUndefined();
    expect(chart.series[0]!.type).toBeUndefined();
    expect(chart.series[0]!.axis).toBeUndefined();
    expect(chart.series[0]!.dataLabels).toBeUndefined();
    expect(chart.series[0]!.markers).toBeUndefined();
  });

  it("keeps the tornado marker and every value", () => {
    const decoded = decodeShipped(SHIPPED_TORNADO);
    expect(decoded.chart).toMatchObject({ kind: "bar", overlap: true });
    expect(decoded.chart!.series[0]!.values).toEqual([-40, 40]);
  });

  it("passes the validator that now allows gaps", () => {
    expect(isChartData(JSON.parse(SHIPPED_STACKED).chart)).toBe(true);
    expect(isChartData(JSON.parse(SHIPPED_TORNADO).chart)).toBe(true);
  });
});
