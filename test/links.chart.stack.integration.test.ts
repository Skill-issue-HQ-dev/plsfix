// A stacked area exported from Excel against the strict Excel fake: a value
// under zero in a stack (also 100 %) puts the approximation sentence on the
// export row, the same one the PowerPoint row uses; a stack from zero up and a
// plain area with a negative value say nothing.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { STACK_NEGATIVE } from "../src/link/chart-model";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["Q1", "Q2", "Q3"];

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost();
});
afterEach(() => uninstallFakeHost());

const series = (name: string, values: number[]) => ({
  name,
  categories: NAMES,
  values,
});

describe("exportActiveChart: a stacked area with a value under zero", () => {
  it("says the slide draws it approximately", async () => {
    const { payload, result } = await exportChart(boot, "AreaStacked", [
      series("A", [10, 20, 30]),
      series("B", [5, -5, 5]),
    ]);
    expect(payload.chart?.kind).toBe("stackedArea");
    expect(result.note).toBe(STACK_NEGATIVE);
    expect(result.note).toBe(
      "values under zero in a stack are drawn approximately",
    );
  });

  it("says it of the 100 % stack too", async () => {
    const { payload, result } = await exportChart(boot, "AreaStacked100", [
      series("A", [10, 20, 30]),
      series("B", [5, -5, 5]),
    ]);
    expect(payload.chart?.percent).toBe(true);
    expect(result.note).toBe(STACK_NEGATIVE);
  });

  it("says nothing of a stack from zero up, nor of a plain area with a negative value", async () => {
    const stack = await exportChart(boot, "AreaStacked", [
      series("A", [10, 20, 30]),
      series("B", [5, 0, 5]),
    ]);
    expect(stack.result.note).toBeUndefined();
    const plain = await exportChart(boot, "Area", [series("A", [10, -20, 30])]);
    expect(plain.payload.chart?.kind).toBe("area");
    expect(plain.result.note).toBeUndefined();
  });
});
