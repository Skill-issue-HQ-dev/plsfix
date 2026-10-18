// lineSeries with a scale for each series, as a combo gives it: the same value
// lands at the height its own series' scale says, a series the scale function
// answers null for is not a line of this chart (nothing drawn, its labels not
// counted in how far apart the others sit), and a plain scale is every series'
// as it always was.

import { describe, expect, it } from "vitest";
import { lineSeries } from "./chart-shapes-line";
import { valueY, type Ellipse, type ValueScale } from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData } from "./link/chart-model";

const PLOT: Box = { left: 0, top: 0, width: 400, height: 100 };

function chartOf(widest: string): ChartData {
  const series = (name: string, values: number[], labels: string[]) => ({
    name,
    values,
    labels,
    colors: values.map(() => "#2EC4B6"),
  });
  return {
    v: 1,
    kind: "line",
    title: null,
    categories: ["A", "B", "C", "D"],
    series: [
      series("low", [5, 10, 5, 10], ["5", "10", "5", "10"]),
      series("high", [50, 100, 50, 100], ["50", "100", "50", "100"]),
      series("wide", [1, 2, 3, 4], [widest, widest, widest, widest]),
    ],
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
}

const markers = (out: ReturnType<typeof lineSeries>) =>
  out.filter((one): one is Ellipse => one.kind === "ellipse");

describe("lineSeries with a scale for each series", () => {
  const data = chartOf("1");
  const own = (low: ValueScale, high: ValueScale) => (index: number) =>
    index === 0 ? low : index === 1 ? high : null;

  it("puts the same height of value at the same height on the slide when each scale is its own series' range", () => {
    const out = lineSeries(
      data,
      PLOT,
      own({ min: 0, max: 10 }, { min: 0, max: 100 }),
    );
    const at = (name: string) =>
      markers(out).find((one) => one.name === name)!.box;
    // 10 of 10 and 100 of 100 are both the top of the plot's value range.
    expect(at("marker 0.1").top + at("marker 0.1").height / 2).toBeCloseTo(
      at("marker 1.1").top + at("marker 1.1").height / 2,
      9,
    );
    // 5 of 10 and 50 of 100 are both half way.
    expect(at("marker 0.0").top).toBeCloseTo(at("marker 1.0").top, 9);
  });

  it("puts each point where valueY says for its own series' scale", () => {
    const low: ValueScale = { min: 0, max: 20 };
    const high: ValueScale = { min: 0, max: 200 };
    const out = lineSeries(data, PLOT, own(low, high));
    const first = markers(out).find((one) => one.name === "marker 1.0")!;
    expect(first.box.top + first.box.height / 2).toBeCloseTo(
      valueY(50, high, PLOT),
      9,
    );
  });

  it("draws nothing for a series the function answers null for", () => {
    const out = lineSeries(
      data,
      PLOT,
      own({ min: 0, max: 10 }, { min: 0, max: 100 }),
    );
    expect(out.some((one) => /^(line|marker|label) 2\./.test(one.name))).toBe(
      false,
    );
  });

  it("does not let a series that is not drawn thin the labels of the others", () => {
    const wide = chartOf("1 000 000 000 000");
    const drawn = lineSeries(wide, PLOT, (index) =>
      index === 0 ? { min: 0, max: 10 } : null,
    );
    const labelNames = drawn
      .filter((one) => one.name.startsWith("label "))
      .map((one) => one.name);
    // Its own "5" and "10" are narrow: every point is labelled.
    expect(labelNames).toEqual([
      "label 0.0",
      "label 0.1",
      "label 0.2",
      "label 0.3",
    ]);
  });

  it("draws every series on one plain scale, as it always did", () => {
    const scale: ValueScale = { min: 0, max: 100 };
    const plain = lineSeries(data, PLOT, scale);
    const asFunction = lineSeries(data, PLOT, () => scale);
    expect(asFunction).toEqual(plain);
  });
});
