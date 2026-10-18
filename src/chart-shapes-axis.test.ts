// Edge to edge: with the category axis on its ticks (ChartData.onTicks) the
// first point of a line or an area is on the plot's left edge and the last on
// its right, n - 1 steps apart, and the names sit under their ticks with the
// two end ones pulled inside; without the flag a point is at the centre of its
// slot, as before. A column or a bar never reads the flag.

import { describe, expect, it } from "vitest";
import {
  BOX,
  DARK,
  areaOf,
  fillRects,
  frame,
  named,
  texts,
} from "../test/chart-area.support";
import { layoutChart, type Line, type Primitive } from "./chart-shapes";
import type { ChartData } from "./link/chart-model";

const area = (extra: Partial<ChartData> = {}): ChartData =>
  areaOf("area", [{ values: [10, 30, 20, 40], color: DARK }], extra);
const line = (extra: Partial<ChartData> = {}): ChartData => ({
  ...area(extra),
  kind: "line",
});

// Where the first point and the last one are drawn: an area's first rectangle
// and last stretch, a line's first and last connector.
function ends(out: Primitive[], kind: "area" | "line"): [number, number] {
  if (kind === "area") {
    const rects = fillRects(out).filter((one) => one.name.startsWith("area "));
    const last = rects.at(-1)!.box;
    return [rects[0]!.box.left, last.left + last.width];
  }
  const joins = named(out, "line ") as Line[];
  const last = joins.at(-1)!.box;
  return [joins[0]!.box.left, last.left + last.width];
}

describe.each(["area", "line"] as const)(
  "a %s with its axis on the ticks",
  (kind) => {
    const make = kind === "area" ? area : line;
    const { plot, slot } = frame(make(), { min: 0, max: 40 });

    it("sits edge to edge: the first point on the plot's left edge, the last on its right", () => {
      const [first, last] = ends(
        layoutChart(make({ onTicks: true }), BOX),
        kind,
      );
      expect(first).toBeCloseTo(plot.left);
      expect(last).toBeCloseTo(plot.left + plot.width);
    });

    it("keeps the centre of the slots without the flag", () => {
      const [first, last] = ends(layoutChart(make(), BOX), kind);
      expect(first).toBeCloseTo(plot.left + slot / 2);
      expect(last).toBeCloseTo(plot.left + plot.width - slot / 2);
    });
  },
);

describe("the category names of an axis on its ticks", () => {
  const names = (data: ChartData) =>
    texts(layoutChart(data, BOX)).filter((one) =>
      one.name.startsWith("category "),
    );
  const { plot } = frame(area(), { min: 0, max: 40 });

  it("are centred on their ticks, the first and the last pulled inside the plot", () => {
    const shown = names(area({ onTicks: true }));
    const step = plot.width / 3;
    const centres = shown.map((one) => one.box.left + one.box.width / 2);
    expect(shown.map((one) => one.align)).toEqual(["l", "c", "c", "r"]);
    expect(centres[1]).toBeCloseTo(plot.left + step);
    expect(centres[2]).toBeCloseTo(plot.left + 2 * step);
    expect(shown[0]!.box.left).toBeCloseTo(plot.left);
    const end = shown.at(-1)!.box;
    expect(end.left + end.width).toBeCloseTo(plot.left + plot.width);
  });

  it("are the slots' own without the flag", () => {
    expect(names(area()).map((one) => one.align)).toEqual(["c", "c", "c", "c"]);
  });
});

describe("a column or a bar chart", () => {
  it("ignores the flag", () => {
    for (const kind of ["column", "bar"] as const) {
      const plain: ChartData = { ...area(), kind };
      const flagged: ChartData = { ...plain, onTicks: true };
      expect(layoutChart(flagged, BOX)).toEqual(layoutChart(plain, BOX));
    }
  });
});
