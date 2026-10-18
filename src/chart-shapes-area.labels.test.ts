// The value labels of a plain area and the area at the caps: a label for each
// point above the line while the series has data labels, else the last point
// alone; under the line where the line touches the top of the plot, below a
// value under zero, thinned when the slots are narrow, always inside the plot;
// and 12 series of 120 points inside the chart box with a name each.

import { describe, expect, it } from "vitest";
import {
  LABEL_HEIGHT,
  LABEL_SIZE,
  bands,
  textWidth,
  type Text,
} from "./chart-shapes";
import {
  BOX,
  DARK,
  INK,
  WHITE,
  areaOf,
  frame,
  insideBox,
  named,
  type Seed,
} from "../test/chart-area.support";
import { layoutArea } from "./chart-shapes-area";

const shown = (out: ReturnType<typeof layoutArea>): Text[] =>
  named(out, "label ") as Text[];

describe("layoutArea: value labels", () => {
  const values = [10, 30, 20, 40];

  it("labels every point above the line while the series has data labels", () => {
    const data = areaOf("area", [{ values }]);
    const labels = shown(layoutArea(data, BOX));
    expect(labels.map((one) => one.text)).toEqual(["10", "30", "20", "40"]);
    const at = frame(data, { min: 0, max: 40 });
    // The first sits right above the line at the first point, centred on it.
    expect(labels[0]!.box.top + LABEL_HEIGHT).toBeCloseTo(at.y(10));
    expect(labels[0]!.box.left + labels[0]!.box.width / 2).toBeCloseTo(at.x(0));
    expect(labels[0]!.color).toBe(INK);
  });

  it("labels only the last point of a series that has no data labels", () => {
    const data = areaOf("area", [{ values, extra: { dataLabels: false } }]);
    const labels = shown(layoutArea(data, BOX));
    expect(labels.map((one) => [one.name, one.text])).toEqual([
      ["label 0.3", "40"],
    ]);
  });

  it("stacks the end labels of several areas apart where they end close together", () => {
    const data = areaOf("area", [
      { values: [200, 250, 288], extra: { dataLabels: false } },
      { values: [100, 200, 281], extra: { dataLabels: false } },
      { values: [100, 200, 285], extra: { dataLabels: false } },
    ]);
    const { plot } = bands(data, BOX);
    const ends = shown(layoutArea(data, BOX)).sort(
      (a, b) => a.box.top - b.box.top,
    );
    expect(ends).toHaveLength(3);
    expect(ends[1]!.box.top - ends[0]!.box.top).toBeGreaterThanOrEqual(
      LABEL_HEIGHT - 1e-9,
    );
    expect(ends[2]!.box.top - ends[1]!.box.top).toBeGreaterThanOrEqual(
      LABEL_HEIGHT - 1e-9,
    );
    // The area that ends highest keeps the highest label.
    expect(ends.map((one) => one.name)).toEqual([
      "label 0.2",
      "label 2.2",
      "label 1.2",
    ]);
    for (const one of ends) expect(insideBox(one, plot, 0.001)).toBe(true);
  });

  it("leaves the end labels of a stack in their bands, where they cannot meet", () => {
    const quiet = { dataLabels: false };
    const data = areaOf("stackedArea", [
      { values: [100, 100, 100], extra: quiet },
      { values: [100, 100, 100], extra: quiet },
    ]);
    const at = frame(data, { min: 0, max: 200 });
    const middles = shown(layoutArea(data, BOX)).map(
      (one) => one.box.top + one.box.height / 2,
    );
    expect(middles).toEqual([
      expect.closeTo((at.y(0) + at.y(100)) / 2, 6),
      expect.closeTo((at.y(100) + at.y(200)) / 2, 6),
    ]);
  });

  it("takes the last point that holds a value when the series ends in blanks", () => {
    const data = areaOf("area", [
      { values: [10, 30, 20, null, null], extra: { dataLabels: false } },
    ]);
    expect(shown(layoutArea(data, BOX)).map((one) => one.name)).toEqual([
      "label 0.2",
    ]);
  });

  it("writes no label for a blank point, which is drawn at zero", () => {
    const data = areaOf("area", [{ values: [10, null, 30] }]);
    expect(shown(layoutArea(data, BOX)).map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.2",
    ]);
  });

  it("decides per series: one with labels beside one without", () => {
    const data = areaOf("area", [
      { values },
      { values, extra: { dataLabels: false } },
    ]);
    const names = shown(layoutArea(data, BOX)).map((one) => one.name);
    expect(names).toEqual([
      "label 0.0",
      "label 0.1",
      "label 0.2",
      "label 0.3",
      "label 1.3",
    ]);
  });

  it("writes no label for a point whose text is blank", () => {
    const data = areaOf("area", [{ values }]);
    data.series[0]!.labels[1] = "";
    expect(shown(layoutArea(data, BOX)).map((one) => one.name)).toEqual([
      "label 0.0",
      "label 0.2",
      "label 0.3",
    ]);
  });

  it("puts the label under a line that touches the top of the plot, on the fill", () => {
    const data = areaOf("area", [{ values: [10, 40, 20, 40], color: DARK }]);
    const { plot } = bands(data, BOX);
    const peak = shown(layoutArea(data, BOX)).find((one) => one.text === "40")!;
    expect(peak.box.top).toBeCloseTo(plot.top);
    expect(peak.color).toBe(WHITE);
  });

  it("puts the label of a value under zero below the line, and inside it at the foot of the plot", () => {
    const data = areaOf("area", [{ values: [20, -5, -10, 10], color: DARK }]);
    const labels = shown(layoutArea(data, BOX));
    const at = frame(data, { min: -10, max: 20 });
    const label = (text: string) => labels.find((one) => one.text === text)!;
    expect(label("-5").box.top).toBeCloseTo(at.y(-5));
    expect(label("-5").color).toBe(INK);
    expect(label("10").box.top + LABEL_HEIGHT).toBeCloseTo(at.y(10));
    // The deepest point is on the foot of the plot: no room below, so the
    // label goes inside the fill, above the line, white on a dark fill.
    expect(label("-10").box.top + LABEL_HEIGHT).toBeCloseTo(at.y(-10));
    expect(label("-10").color).toBe(WHITE);
  });

  it("thins the labels to every n-th once a slot is narrower than the widest", () => {
    const many = Array.from({ length: 60 }, (_one, i) => 1000 + i);
    const data = areaOf("area", [{ values: many }]);
    const labels = shown(layoutArea(data, BOX));
    const slot = bands(data, BOX).plot.width / 60;
    const stride = Math.ceil(textWidth("1000", LABEL_SIZE) / slot);
    expect(stride).toBeGreaterThan(1);
    expect(labels.length).toBe(Math.ceil(60 / stride));
    expect(labels[0]!.name).toBe("label 0.0");
    expect(labels[1]!.name).toBe(`label 0.${String(stride)}`);
  });

  it("keeps every label inside the plot, even at its two ends", () => {
    const wide = Array.from({ length: 30 }, (_one, i) => 25_000_000 + i);
    const data = areaOf("area", [{ values: wide }]);
    const { plot } = bands(data, BOX);
    const labels = shown(layoutArea(data, BOX));
    expect(labels.length).toBeGreaterThan(1);
    for (const one of labels) expect(insideBox(one, plot, 0.001)).toBe(true);
  });
});

describe("layoutArea at the caps", () => {
  const points = 120;
  const seeds = (count: number): Seed[] =>
    Array.from({ length: count }, (_one, j) => ({
      values: Array.from(
        { length: points },
        (_two, i) => ((i * 7 + j * 13) % 50) - 10,
      ),
    }));

  it.each(["area", "stackedArea"] as const)(
    "keeps every %s shape of 12 series and 120 points inside the chart box, with sides of its own and a name of its own",
    (kind) => {
      const out = layoutArea(areaOf(kind, seeds(12)), BOX);
      expect(out.length).toBeGreaterThan(12 * 119);
      expect(out.filter((one) => !insideBox(one, BOX))).toEqual([]);
      const names = out.map((one) => one.name);
      expect(new Set(names).size).toBe(names.length);
    },
  );
});
