// What the series reads make of a dimension's raw cells: a blank cell (an empty
// string, nothing, or #N/A) is a gap, never a zero. Everything that talks to
// office.js is covered by test/links.chart.integration.test.ts against the
// strict fake and test/hunt/excel.link-chart-read.hunt.test.ts against a
// hand-built context.

import { describe, expect, it } from "vitest";
import { valuesOf } from "./chart-read-series";

describe("valuesOf", () => {
  it("reads the numbers a dimension gives back, whatever their type", () => {
    expect(valuesOf(["1", 2.5])).toEqual([1, 2.5]);
    expect(valuesOf([])).toEqual([]);
  });

  it("keeps a blank cell as a gap, however the host spells it", () => {
    expect(valuesOf([1, "", 3])).toEqual([1, null, 3]);
    expect(valuesOf(["1", "  "])).toEqual([1, null]);
    expect(valuesOf([1, null, 3])).toEqual([1, null, 3]);
    expect(valuesOf([undefined, 2])).toEqual([null, 2]);
  });

  it("keeps #N/A as a gap, the way Excel plots it", () => {
    expect(valuesOf([1, "#N/A", 3])).toEqual([1, null, 3]);
  });

  it("keeps a series of nothing but gaps", () => {
    expect(valuesOf(["", null])).toEqual([null, null]);
  });

  it("refuses a dimension that holds anything else that is not a number", () => {
    expect(valuesOf(["x"])).toBeNull();
    expect(valuesOf([1, "n/a"])).toBeNull();
    expect(valuesOf([1, "#DIV/0!"])).toBeNull();
    expect(valuesOf([Number.POSITIVE_INFINITY])).toBeNull();
  });
});
