// The combo layout, geometry and scales: the column series (clustered, or
// stacked) on the primary scale exactly as the column layout draws them, the
// line series as the line layout draws them, each series on ITS axis - a series
// on the secondary axis gets a scale of its own from its own values, zero
// included, so a 0-40 % margin line over columns of EUR 0-13 392 k uses the
// plot's full height. A combo with a series the slide cannot draw is null.
// Pure geometry in slide points; labels and legend are in
// chart-shapes-combo.labels.test.ts.

import { describe, expect, it } from "vitest";
import { BAR_FILL } from "./chart-shapes";
import { layoutCombo } from "./chart-shapes-combo";
import {
  BOX,
  bars,
  centreOf,
  comboOf,
  frame,
  markers,
  named,
  segments,
} from "../test/chart-combo.support";

const REVENUE = [1020, 985, 1075, 1090];
const MARGIN = [0.1284, 0.1178, 0.146, 0.155];

describe("layoutCombo: columns and a line on the secondary axis", () => {
  const data = comboOf([
    { type: "column", values: REVENUE },
    { type: "line", axis: "secondary", values: MARGIN },
  ]);
  const out = layoutCombo(data, BOX)!;
  const f = frame(data);

  it("draws a bar for each column value, as the column layout does: 0.6 of its slot, centred", () => {
    const drawn = bars(out);
    expect(drawn).toHaveLength(4);
    drawn.forEach((bar, i) => {
      expect(bar.name).toBe(`bar 0.${String(i)}`);
      expect(bar.box.width).toBeCloseTo(BAR_FILL * f.slot, 9);
      expect(bar.box.left + bar.box.width / 2).toBeCloseTo(f.x(i), 9);
    });
  });

  it("scales the columns by their own values: the tallest stands a label's room under the plot's top, every bar on the zero line", () => {
    const drawn = bars(out);
    drawn.forEach((bar, i) => {
      expect(bar.box.top).toBeCloseTo(f.y(REVENUE[i]!, 1090), 9);
      expect(bar.box.top + bar.box.height).toBeCloseTo(f.y(0, 1090), 9);
    });
    expect(drawn[3]!.box.top).toBeCloseTo(f.plot.top + f.room, 9);
  });

  it("scales the line by its own values, not by the columns': its highest point is as high as the tallest column", () => {
    const top = centreOf(named(out, "marker 1.3"));
    expect(top.y).toBeCloseTo(f.y(0.155, 0.155), 9);
    expect(top.y).toBeCloseTo(bars(out)[3]!.box.top, 9);
    expect(top.x).toBeCloseTo(f.x(3), 9);
  });

  it("puts a lower value where its own scale says, with zero in the scale", () => {
    const point = centreOf(named(out, "marker 1.1"));
    expect(point.y).toBeCloseTo(f.y(0.1178, 0.155), 9);
    // Zero is in the scale, so 0.1178 is three quarters of the way up, not at the floor.
    expect(point.y).toBeLessThan(f.y(0, 0.155) - 100);
  });

  it("joins the points with the line layout's segments, none into a point that is not there", () => {
    expect(segments(out).map((one) => one.name)).toEqual([
      "line 1.0",
      "line 1.1",
      "line 1.2",
    ]);
  });

  it("draws the bars first, then the zero line, then the line over them", () => {
    const at = (name: string) => out.findIndex((one) => one.name === name);
    expect(at("bar 0.3")).toBeLessThan(at("baseline"));
    expect(at("baseline")).toBeLessThan(at("line 1.0"));
    expect(at("baseline")).toBeLessThan(at("marker 1.0"));
    expect(named(out, "baseline").box.top).toBeCloseTo(f.y(0, 1090), 9);
  });
});

describe("layoutCombo: a margin line over big columns uses the plot's full height", () => {
  const data = comboOf([
    { type: "column", values: [13392, 10000, 12000] },
    { type: "line", axis: "secondary", values: [0, 0.2, 0.4] },
  ]);
  const out = layoutCombo(data, BOX)!;
  const f = frame(data);

  it("takes the margin's highest point to the top of the value area and its middle to half way", () => {
    expect(centreOf(named(out, "marker 1.2")).y).toBeCloseTo(f.y(0.4, 0.4), 9);
    expect(centreOf(named(out, "marker 1.1")).y).toBeCloseTo(f.y(0.2, 0.4), 9);
    // The columns, a thousand times the line's numbers, did not flatten it.
    const height = f.plot.height - f.room;
    expect(
      centreOf(named(out, "marker 1.1")).y - (f.plot.top + f.room),
    ).toBeCloseTo(height / 2, 9);
  });

  it("gives the column its own full height beside it", () => {
    expect(bars(out)[0]!.box.top).toBeCloseTo(f.plot.top + f.room, 9);
  });
});

describe("layoutCombo: a line on the primary axis shares the columns' scale", () => {
  it("puts a value of the line level with a column of the same value", () => {
    const data = comboOf([
      { type: "column", values: [100, 300, 200] },
      { type: "line", values: [300, 100, 200] },
    ]);
    const out = layoutCombo(data, BOX)!;
    expect(centreOf(named(out, "marker 1.0")).y).toBeCloseTo(
      named(out, "bar 0.1").box.top,
      9,
    );
    expect(centreOf(named(out, "marker 1.1")).y).toBeCloseTo(
      named(out, "bar 0.0").box.top,
      9,
    );
  });

  it("lets a line taller than every column set the scale for both", () => {
    const data = comboOf([
      { type: "column", values: [100, 200] },
      { type: "line", values: [400, 300] },
    ]);
    const out = layoutCombo(data, BOX)!;
    const f = frame(data);
    expect(centreOf(named(out, "marker 1.0")).y).toBeCloseTo(f.y(400, 400), 9);
    expect(named(out, "bar 0.1").box.top).toBeCloseTo(f.y(200, 400), 9);
  });
});

describe("layoutCombo: clustered columns", () => {
  it("shares each slot among the column series only, never counting a line", () => {
    const data = comboOf([
      { type: "column", values: [10, 20, 30] },
      { type: "line", axis: "secondary", values: [1, 2, 3] },
      { type: "column", values: [30, 20, 10] },
    ]);
    const out = layoutCombo(data, BOX)!;
    const f = frame(data);
    const first = named(out, "bar 0.1");
    const second = named(out, "bar 2.1");
    expect(first.box.width).toBeCloseTo((BAR_FILL * f.slot) / 2, 9);
    expect(second.box.width).toBeCloseTo((BAR_FILL * f.slot) / 2, 9);
    // The series side by side in the slot, the first at the left.
    expect(second.box.left).toBeCloseTo(first.box.left + first.box.width, 9);
    expect(
      (first.box.left + second.box.left + second.box.width) / 2,
    ).toBeCloseTo(f.x(1), 9);
  });

  it("names a series' shapes by its place in the payload, whatever it is drawn as", () => {
    const data = comboOf([
      { type: "line", axis: "secondary", values: [1, 2] },
      { type: "column", values: [5, 6] },
      { type: "line", values: [4, 5] },
    ]);
    const names = layoutCombo(data, BOX)!.map((one) => one.name);
    expect(names).toContain("bar 1.0");
    expect(names).toContain("marker 0.0");
    expect(names).toContain("marker 2.1");
    expect(names).toContain("line 0.0");
    expect(names).not.toContain("bar 0.0");
  });

  it("draws no bar for a blank cell and breaks the line at one", () => {
    const data = comboOf([
      { type: "column", values: [10, null, 30, 40] },
      { type: "line", axis: "secondary", values: [1, 2, null, 4] },
    ]);
    const out = layoutCombo(data, BOX)!;
    expect(bars(out).map((one) => one.name)).toEqual([
      "bar 0.0",
      "bar 0.2",
      "bar 0.3",
    ]);
    expect(markers(out).map((one) => one.name)).toEqual([
      "marker 1.0",
      "marker 1.1",
      "marker 1.3",
    ]);
    expect(segments(out).map((one) => one.name)).toEqual(["line 1.0"]);
  });
});

describe("layoutCombo: stacked columns", () => {
  const data = comboOf([
    { type: "stackedColumn", values: [100, 200, 150] },
    { type: "stackedColumn", values: [50, 50, 100] },
    { type: "line", axis: "secondary", values: [0.1, 0.2, 0.3] },
  ]);
  const out = layoutCombo(data, BOX)!;
  const f = frame(data);

  it("piles the series up in each column, the second on the first, on a scale of the stacks' sums", () => {
    for (let i = 0; i < 3; i += 1) {
      const lower = named(out, `bar 0.${String(i)}`);
      const upper = named(out, `bar 1.${String(i)}`);
      expect(upper.box.top + upper.box.height).toBeCloseTo(lower.box.top, 9);
      // One bar the width of the slot's share, not split among the series.
      expect(lower.box.width).toBeCloseTo(BAR_FILL * f.slot, 9);
      expect(upper.box.left).toBeCloseTo(lower.box.left, 9);
    }
    // 250 is the tallest stack (the second and third are 250): its top is at the value area's top.
    expect(named(out, "bar 1.1").box.top).toBeCloseTo(f.y(250, 250), 9);
    expect(named(out, "bar 0.1").box.top).toBeCloseTo(f.y(200, 250), 9);
  });

  it("keeps the line on its own scale beside the stacks", () => {
    expect(centreOf(named(out, "marker 2.2")).y).toBeCloseTo(f.y(0.3, 0.3), 9);
  });

  it("lets a line on the primary axis raise the scale of the stacks", () => {
    const raised = comboOf([
      { type: "stackedColumn", values: [100, 200] },
      { type: "stackedColumn", values: [50, 50] },
      { type: "line", values: [500, 400] },
    ]);
    const drawn = layoutCombo(raised, BOX)!;
    const g = frame(raised);
    expect(named(drawn, "bar 1.1").box.top).toBeCloseTo(g.y(250, 500), 9);
    expect(centreOf(named(drawn, "marker 2.0")).y).toBeCloseTo(
      g.y(500, 500),
      9,
    );
  });
});

describe("layoutCombo: values under zero", () => {
  it("draws a negative column down from the zero line, which sits where the primary scale's zero is", () => {
    const data = comboOf([
      { type: "column", values: [-50, 100] },
      { type: "line", axis: "secondary", values: [0.1, 0.2] },
    ]);
    const out = layoutCombo(data, BOX)!;
    const f = frame(data);
    const zero = f.y(0, 100, -50);
    expect(named(out, "baseline").box.top).toBeCloseTo(zero, 9);
    const down = named(out, "bar 0.0");
    expect(down.box.top).toBeCloseTo(zero, 9);
    expect(down.box.top + down.box.height).toBeCloseTo(f.y(-50, 100, -50), 9);
    expect(named(out, "bar 0.1").box.top).toBeCloseTo(f.y(100, 100, -50), 9);
  });
});

describe("layoutCombo: every series on the secondary axis is the one axis", () => {
  it("draws columns and a line all marked secondary on one shared scale, like a primary pair", () => {
    const data = comboOf([
      { type: "column", axis: "secondary", values: [100, 300] },
      { type: "line", axis: "secondary", values: [300, 100] },
    ]);
    const out = layoutCombo(data, BOX)!;
    expect(out).not.toBeNull();
    expect(centreOf(named(out, "marker 1.0")).y).toBeCloseTo(
      named(out, "bar 0.1").box.top,
      9,
    );
  });
});

describe("layoutCombo: a combo with a series the slide cannot draw", () => {
  it.each(["bar", "area", "stackedArea", "scatter"] as const)(
    "is null for a %s series, and for columns on the secondary axis",
    (type) => {
      const data = comboOf([
        { type: "column", values: [1, 2] },
        { type, values: [3, 4] },
      ]);
      expect(layoutCombo(data, BOX)).toBeNull();
    },
  );

  it("is null for columns on the secondary axis beside a primary line", () => {
    const data = comboOf([
      { type: "column", axis: "secondary", values: [1, 2] },
      { type: "line", values: [3, 4] },
    ]);
    expect(layoutCombo(data, BOX)).toBeNull();
  });
});
