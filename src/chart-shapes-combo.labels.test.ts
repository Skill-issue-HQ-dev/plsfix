// The combo layout, labels and frame: a column's value above its bar and a
// line's above its marker, each in the text its own series carries (a margin in
// its percent format, never reformatted), never on the title; a series without
// data labels keeps the line layout's last value; the category names under the
// plot; the legend of every series in the payload's (Excel's) order; and every
// primitive inside its box with a name of its own, at the caps 120 / 12.
// Pure geometry in slide points.

import { describe, expect, it } from "vitest";
import { TITLE_BAND } from "./chart-shapes";
import { layoutCombo } from "./chart-shapes-combo";
import {
  BOX,
  BRASS,
  INK,
  NAVY,
  centreOf,
  comboOf,
  frame,
  insideBox,
  labels,
  named,
  texts,
  type Seed,
} from "../test/chart-combo.support";

const MONTHS = ["12.8%", "11.8%", "14.6%", "15.5%"];

describe("layoutCombo: the value labels", () => {
  const data = comboOf([
    { type: "column", values: [1020, 985, 1075, 1090] },
    {
      type: "line",
      axis: "secondary",
      values: [0.1284, 0.1178, 0.146, 0.155],
      labels: MONTHS,
    },
  ]);
  const out = layoutCombo(data, BOX)!;
  const f = frame(data);

  it("writes a column's value above its bar, centred on it", () => {
    const one = named(out, "label 0.2");
    const bar = named(out, "bar 0.2");
    expect(one).toMatchObject({ kind: "text", text: "1075", align: "c" });
    expect(one.box.top + one.box.height).toBeCloseTo(bar.box.top, 9);
    expect(one.box.left + one.box.width / 2).toBeCloseTo(f.x(2), 9);
  });

  it("writes a line's value above its marker in the text its series carries: the margin as its own percent, not the columns' format", () => {
    const [shown] = labels(out).filter((one) => one.name === "label 1.1");
    expect(shown!.text).toBe("11.8%");
    expect(shown!.box.top + shown!.box.height).toBeLessThanOrEqual(
      centreOf(named(out, "marker 1.1")).y,
    );
    expect(
      labels(out).filter((one) => one.name.startsWith("label 1.")),
    ).toHaveLength(4);
  });

  it("never lets a label touch the title band, even over the tallest column and the highest point", () => {
    for (const one of labels(out)) {
      expect(one.box.top, one.name).toBeGreaterThanOrEqual(
        BOX.top + TITLE_BAND - 1e-6,
      );
    }
    // The tallest column's label is the lowest it can stand: just under the band.
    expect(named(out, "label 0.3").box.top).toBeGreaterThanOrEqual(
      f.plot.top - 1e-6,
    );
  });

  it("keeps a series without data labels to its last value, as the line layout does", () => {
    const quiet = comboOf([
      { type: "column", values: [10, 20, 30], extra: { dataLabels: false } },
      {
        type: "line",
        axis: "secondary",
        values: [1, 2, 3],
        extra: { dataLabels: false },
      },
    ]);
    const names = layoutCombo(quiet, BOX)!
      .filter((one) => one.name.startsWith("label "))
      .map((one) => one.name);
    expect(names).toEqual(["label 1.2"]);
  });

  it("writes the columns' labels in the ink of the chart", () => {
    expect(named(out, "label 0.0")).toMatchObject({ color: INK });
  });
});

describe("layoutCombo: the categories and the legend", () => {
  const data = comboOf(
    [
      { type: "column", values: [1, 2, 3], color: BRASS },
      { type: "line", axis: "secondary", values: [4, 5, 6], color: NAVY },
    ],
    { title: "Revenue and margin" },
  );
  const out = layoutCombo(data, BOX)!;

  it("writes the title first and the category names under the plot", () => {
    expect(out[0]).toMatchObject({ name: "title", text: "Revenue and margin" });
    const categories = texts(out).filter((one) =>
      one.name.startsWith("category "),
    );
    expect(categories.map((one) => one.text)).toEqual(["C0", "C1", "C2"]);
    const { plot } = frame(data);
    for (const one of categories) {
      expect(one.box.top).toBeGreaterThanOrEqual(plot.top + plot.height - 1e-6);
    }
  });

  it("lists every series in the payload's order, each swatch in its series' colour", () => {
    const words = texts(out)
      .filter((one) => one.name.startsWith("legend "))
      .map((one) => one.text);
    expect(words).toEqual(["s0", "s1"]);
    const swatches = out.filter((one) => one.name.startsWith("legend swatch"));
    expect(
      swatches.map((one) => (one.kind === "rect" ? one.color : "")),
    ).toEqual([BRASS, NAVY]);
  });
});

describe("layoutCombo: inside its box, at the caps", () => {
  const BLANKS = 7;

  function many(columns: number, lines: number, points: number): Seed[] {
    const seeds: Seed[] = [];
    for (let j = 0; j < columns + lines; j += 1) {
      const values = Array.from({ length: points }, (_one, i) =>
        i % BLANKS === j % BLANKS ? null : ((i * 13 + j * 7) % 90) - 20,
      );
      seeds.push(
        j < columns
          ? { type: "column", values }
          : {
              type: "line",
              axis: j % 2 === 0 ? "secondary" : "primary",
              values,
            },
      );
    }
    return seeds;
  }

  it.each([
    [6, 6, 120],
    [1, 11, 120],
    [11, 1, 120],
    [3, 3, 2],
    [2, 2, 24],
  ])(
    "keeps every primitive in the box with its own name: %i columns, %i lines, %i points",
    (columns, lines, points) => {
      const data = comboOf(many(columns, lines, points), { title: "Caps" });
      const out = layoutCombo(data, BOX)!;
      expect(out.length).toBeGreaterThan(0);
      expect(out.filter((one) => !insideBox(one, BOX))).toEqual([]);
      expect(new Set(out.map((one) => one.name)).size).toBe(out.length);
    },
  );

  it("holds in a box as small as the pane draws", () => {
    const small = { left: 0, top: 0, width: 200, height: 120 };
    const data = comboOf(many(2, 2, 12), { title: "Small" });
    const out = layoutCombo(data, small)!;
    expect(out.filter((one) => !insideBox(one, small))).toEqual([]);
  });
});
