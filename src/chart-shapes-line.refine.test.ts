// The line kind's refinements (C3): markers only for a series that has them and
// never past 24 points, the value of every point while the series has data
// labels and the last point alone when it has none, the last point being the
// last cell that holds a value, a smooth line drawn exactly as a straight one,
// and each connector flagged for the way it climbs so the slide can draw a
// rising one with LineInverse and a falling or flat one as a straight line.
// Called on its own with a hand-set plot and scale.

import { describe, expect, it } from "vitest";
import { lineSeries } from "./chart-shapes-line";
import {
  LABEL_HEIGHT,
  TITLE_BAND,
  type Ellipse,
  type Line,
  type Primitive,
  type Text,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";

const PLOT: Box = { left: 0, top: 500, width: 400, height: 100 };
const SCALE: ValueScale = { min: 0, max: 20 };
const CATEGORIES = ["A", "B", "C", "D"];

function lineData(
  series: Partial<ChartSeries>[],
  categories = CATEGORIES,
): ChartData {
  return {
    v: 1,
    kind: "line",
    title: null,
    categories,
    series: series.map((one, j): ChartSeries => {
      const values = one.values ?? [10, 5, 15, 10];
      return {
        name: `s${String(j)}`,
        values,
        labels: values.map((value) => (value === null ? "" : String(value))),
        colors: values.map(() => "#2EC4B6"),
        ...one,
      };
    }),
    font: "Arial",
    ink: "#111111",
    titleColor: "#000000",
  };
}

const markers = (out: Primitive[]): Ellipse[] =>
  out.filter((one): one is Ellipse => one.kind === "ellipse");
const lines = (out: Primitive[]): Line[] =>
  out.filter((one): one is Line => one.kind === "line");
const labels = (out: Primitive[]): Text[] =>
  out.filter((one): one is Text => one.kind === "text");
const draw = (data: ChartData): Primitive[] => lineSeries(data, PLOT, SCALE);

describe("a line's markers", () => {
  it("are drawn for a series that has them, and not for one that says it has none", () => {
    const out = draw(lineData([{}, { markers: false }]));
    expect(markers(out).map((one) => one.name)).toEqual([
      "marker 0.0",
      "marker 0.1",
      "marker 0.2",
      "marker 0.3",
    ]);
  });

  it("are all off past 24 points, whatever the series says", () => {
    const points = Array.from({ length: 25 }, (_one, i) => `C${String(i)}`);
    const values = points.map((_one, i) => i % 20);
    const flagged = lineData([{ values, markers: true }], points);
    expect(markers(draw(flagged))).toEqual([]);
    const at24 = lineData(
      [{ values: values.slice(0, 24) }],
      points.slice(0, 24),
    );
    expect(markers(draw(at24))).toHaveLength(24);
  });
});

describe("a line's value labels", () => {
  it("are on every point while the series has data labels", () => {
    const out = draw(lineData([{}]));
    expect(labels(out).map((one) => one.text)).toEqual(["10", "5", "15", "10"]);
  });

  it("are on the last point alone when the series has none", () => {
    const out = draw(lineData([{ dataLabels: false }]));
    expect(labels(out).map((one) => [one.name, one.text])).toEqual([
      ["label 0.3", "10"],
    ]);
  });

  it("are on the last cell that holds a value when the series ends in blanks", () => {
    const out = draw(
      lineData([{ values: [10, 5, 15, null], dataLabels: false }]),
    );
    expect(labels(out).map((one) => one.name)).toEqual(["label 0.2"]);
  });

  it("are on nothing at all when every cell of a series without data labels is blank", () => {
    const out = draw(
      lineData([{ values: [null, null, null, null], dataLabels: false }]),
    );
    expect(out).toEqual([]);
  });

  it("sit above the point of the last label as any other does", () => {
    const out = draw(
      lineData([{ values: [10, 5, 15, 10], dataLabels: false }]),
    );
    const [label] = labels(out);
    const [last] = markers(out).slice(-1);
    expect(label!.box.top + LABEL_HEIGHT).toBeCloseTo(last!.box.top);
  });

  it("keep the last point even with markers off and past 24 points", () => {
    const points = Array.from({ length: 40 }, (_one, i) => `C${String(i)}`);
    const values = points.map((_one, i) => (i % 15) + 1);
    const out = draw(
      lineData([{ values, markers: false, dataLabels: false }], points),
    );
    expect(markers(out)).toEqual([]);
    expect(labels(out).map((one) => one.name)).toEqual(["label 0.39"]);
  });

  it("decide per series, each last point labelled where its series has no labels", () => {
    const out = draw(
      lineData([
        {},
        { dataLabels: false },
        { dataLabels: false, values: [1, 2, 3, null] },
      ]),
    );
    expect(labels(out).map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.1",
      "label 0.2",
      "label 0.3",
      "label 1.3",
      "label 2.2",
    ]);
  });
});

// Several lines that end close together each keep the value of their last
// point, and those labels would sit on one another: they are stacked apart,
// the higher line's label staying the higher.
describe("a line chart's end labels", () => {
  const tall: Box = { left: 0, top: 500, width: 400, height: 200 };
  const wide: ValueScale = { min: 0, max: 400 };
  const quiet = { dataLabels: false };

  function ends(values: number[][], extra: Partial<ChartSeries> = quiet) {
    const data = lineData(
      values.map((one) => ({ values: one, ...extra })),
      CATEGORIES,
    );
    const out = lineSeries(data, tall, wide);
    return labels(out).filter((one) => one.name.endsWith(".3"));
  }

  it("stay over their own markers when the lines end far apart", () => {
    const [low, high] = ends([
      [100, 100, 100, 100],
      [300, 300, 300, 300],
    ]);
    const markerTop = (value: number) => 500 + 200 * (1 - value / 400) - 3; // the marker is 6 points, centred
    expect(low!.box.top + LABEL_HEIGHT).toBeCloseTo(markerTop(100), 0);
    expect(high!.box.top + LABEL_HEIGHT).toBeCloseTo(markerTop(300), 0);
  });

  it("are pushed apart a label's height when the lines end within it of each other", () => {
    const [a, b, c] = ends([
      [288, 288, 288, 288],
      [281, 281, 281, 281],
      [285, 285, 285, 285],
    ]);
    const sorted = [a!, b!, c!].map((one) => one.box.top).sort((x, y) => x - y);
    expect(sorted[1]! - sorted[0]!).toBeGreaterThanOrEqual(LABEL_HEIGHT - 1e-9);
    expect(sorted[2]! - sorted[1]!).toBeGreaterThanOrEqual(LABEL_HEIGHT - 1e-9);
    // The line that ends highest has the highest label.
    expect(a!.box.top).toBeLessThan(c!.box.top);
    expect(c!.box.top).toBeLessThan(b!.box.top);
  });

  it("stay inside the room above and below the plot", () => {
    const squeezed = ends([
      [399, 399, 399, 399],
      [398, 398, 398, 398],
      [1, 1, 1, 1],
      [2, 2, 2, 2],
    ]);
    for (const one of squeezed) {
      expect(one.box.top).toBeGreaterThanOrEqual(tall.top - TITLE_BAND - 1e-9);
      expect(one.box.top + one.box.height).toBeLessThanOrEqual(
        tall.top + tall.height + 1e-9,
      );
    }
  });

  it("are the only labels that move: a series with data labels keeps its own", () => {
    const data = lineData([
      { values: [288, 288, 288, 288] },
      { values: [281, 281, 281, 281], dataLabels: false },
    ]);
    const out = labels(lineSeries(data, tall, wide));
    const own = out.filter((one) => one.name.startsWith("label 0."));
    expect(own).toHaveLength(4);
    const natural = lineSeries(
      lineData([{ values: [288, 288, 288, 288] }]),
      tall,
      wide,
    );
    expect(own.map((one) => one.box)).toEqual(
      labels(natural).map((one) => one.box),
    );
  });

  it("do not move one another when they are over different points", () => {
    // The first line ends a point early, so the two labels are in two columns.
    const data = lineData([
      { values: [288, 288, 288, null], ...quiet },
      { values: [281, 281, 281, 281], ...quiet },
    ]);
    const out = labels(lineSeries(data, tall, wide));
    const alone = labels(lineSeries(lineData([data.series[0]!]), tall, wide));
    expect(out.find((one) => one.name === "label 0.2")!.box).toEqual(
      alone.find((one) => one.name === "label 0.2")!.box,
    );
  });
});

describe("a smooth line", () => {
  it("is drawn exactly as the straight one: smoothing changes nothing in the layout", () => {
    const straight = draw(lineData([{ values: [10, 5, 15, 10] }]));
    const smooth = draw(lineData([{ values: [10, 5, 15, 10], smooth: true }]));
    expect(smooth).toEqual(straight);
    expect(lines(smooth)).toHaveLength(3);
  });
});

describe("a line's connectors", () => {
  it("climb, fall or stay level, and say so with the rising flag alone", () => {
    const out = draw(lineData([{ values: [10, 15, 15, 5] }]));
    const found = lines(out).map((one) => [
      one.name,
      one.rising,
      one.box.height === 0,
    ]);
    expect(found).toEqual([
      ["line 0.0", true, false],
      ["line 0.1", false, true],
      ["line 0.2", false, false],
    ]);
  });

  it("never carry a negative side, whatever way they run", () => {
    const out = draw(lineData([{ values: [20, 0, 20, 0] }]));
    for (const one of lines(out)) {
      expect(one.box.width).toBeGreaterThan(0);
      expect(one.box.height).toBeGreaterThanOrEqual(0);
    }
  });

  it("start where the point before them is, and end on the next", () => {
    const out = draw(lineData([{ values: [10, 15, 15, 5] }]));
    const [first] = lines(out);
    const [one, two] = markers(out);
    expect(first!.box.left).toBeCloseTo(one!.box.left + one!.box.width / 2);
    expect(first!.box.left + first!.box.width).toBeCloseTo(
      two!.box.left + two!.box.width / 2,
    );
  });
});
