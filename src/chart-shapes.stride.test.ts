// keptIndices thins an axis to every k-th label. A stride under 1 - the slot of
// a plot with no width or a negative one, in a box too small to hold a plot -
// must not walk backwards or stand still: it used to push into one array for
// seconds until the engine refused ("Invalid array length"), freezing the pane.

import { describe, expect, it } from "vitest";
import { categoryLabels, keptIndices, thinningFactor } from "./chart-shapes";
import { chartOf } from "../test/hunt/ppt.chart-of.hunt3.support";

describe("keptIndices", () => {
  it("keeps every k-th index, the first always", () => {
    expect(keptIndices(5, 1)).toEqual([0, 1, 2, 3, 4]);
    expect(keptIndices(5, 2)).toEqual([0, 2, 4]);
    expect(keptIndices(5, 7)).toEqual([0]);
    expect(keptIndices(0, 2)).toEqual([]);
  });

  it("keeps only the first label for a stride that cannot advance", () => {
    for (const stride of [0, -1, -3.4, Number.NaN, 0.5, -Infinity]) {
      expect(keptIndices(40, stride), String(stride)).toEqual([0]);
    }
    expect(keptIndices(40, Infinity)).toEqual([0]);
    expect(keptIndices(0, -2)).toEqual([]);
  });

  it("answers at once for the stride a negative slot gives", () => {
    const stride = thinningFactor(-12.5, 40);
    expect(stride).toBeLessThan(0);
    const started = performance.now();
    expect(keptIndices(30, stride)).toEqual([0]);
    expect(performance.now() - started).toBeLessThan(50);
  });
});

describe("categoryLabels in a plot with no height", () => {
  it("returns without hanging: one label at most", () => {
    const data = chartOf("bar", 8, 1, { title: null });
    const plot = { left: 100, top: 40, width: 80, height: -20 };
    const band = { left: 0, top: 40, width: 100, height: -20 };
    const started = performance.now();
    const labels = categoryLabels(data, plot, band, true);
    expect(labels.length).toBeLessThanOrEqual(1);
    expect(performance.now() - started).toBeLessThan(500);
  });
});
