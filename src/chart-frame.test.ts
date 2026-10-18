// The layout box a rebuilt chart is drawn in: derived from where the group
// stands (the union of its shapes) so that the drawing lands exactly there. A
// titled chart fills its box and needs no help; an untitled one leaves a band
// empty above its plot, and a pie or a doughnut is as wide as its widest legend
// row, so many box widths draw the same reach. The union is measured here with
// its own few lines, not with the module's.

import { describe, expect, it } from "vitest";
import { frameFor, type Reach } from "./chart-frame";
import { layoutChart } from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartKind } from "./link/chart-model";
import { chartOf } from "../test/hunt/ppt.chart-of.hunt3.support";

function drawnAt(data: ChartData, box: Box): Reach {
  const boxes = layoutChart(data, { ...box, left: 0, top: 0 }).map(
    (primitive) => primitive.box,
  );
  const left = Math.min(...boxes.map((one) => one.left));
  const right = Math.max(...boxes.map((one) => one.left + one.width));
  const top = Math.min(...boxes.map((one) => one.top));
  return { left: box.left + left, top: box.top + top, width: right - left };
}

function expectReach(got: Reach, want: Reach): void {
  expect(got.left).toBeCloseTo(want.left, 0);
  expect(got.top).toBeCloseTo(want.top, 0);
  expect(got.width).toBeCloseTo(want.width, 0);
}

const ASPECT = 270 / 480;
const KINDS: [ChartKind, number][] = [
  ["column", 1],
  ["column", 3],
  ["bar", 2],
  ["pie", 1],
  ["line", 2],
  ["waterfall", 1],
  ["stackedColumn", 2],
  ["stackedBar", 2],
  ["doughnut", 1],
  ["doughnut", 2],
  ["area", 2],
  ["stackedArea", 3],
  ["combo", 3],
];

function box(width: number, aspect = ASPECT): Box {
  return { left: 123.4, top: 56.7, width, height: Math.round(width * aspect) };
}

describe("a titled chart", () => {
  it("is drawn in the box that is the group's own: its title spans the top edge", () => {
    for (const kind of ["column", "bar", "line", "pie"] as const) {
      const data = chartOf(kind, 6, kind === "pie" ? 1 : 2);
      const at = box(360);
      const reach = drawnAt(data, at);
      const frame = frameFor(data, reach, ASPECT, 360);
      expect(frame.left, kind).toBeCloseTo(at.left, 6);
      expect(frame.top, kind).toBeCloseTo(at.top, 6);
      expect(frame.width, kind).toBeCloseTo(360, 6);
    }
  });
});

describe("an untitled chart", () => {
  it("is drawn above its group by the band its layout leaves empty", () => {
    const column = chartOf("column", 6, 2, { title: null });
    const bar = chartOf("bar", 6, 2, { title: null });
    const there: Reach = { left: 300, top: 200, width: 360 };

    // The title band (18) and, on a column chart, the room its labels keep (3).
    expect(frameFor(column, there, ASPECT, 360).top).toBeCloseTo(200 - 21, 6);
    expect(frameFor(bar, there, ASPECT, 360).top).toBeCloseTo(200 - 18, 6);
    expect(frameFor(column, there, ASPECT, 360).left).toBeCloseTo(300, 6);
    expect(frameFor(column, there, ASPECT, 360).width).toBeCloseTo(360, 6);
  });

  it("gets its height from the aspect at the width it keeps", () => {
    const column = chartOf("column", 6, 2, { title: null });
    const frame = frameFor(column, { left: 0, top: 0, width: 400 }, 0.5, 400);
    expect(frame.height).toBe(200);
  });
});

describe("every kind, over sizes and aspects", () => {
  it("lands the drawing exactly where the old one stood, and keeps the width it was drawn at", () => {
    for (const [kind, series] of KINDS) {
      for (const points of [3, 6, 12]) {
        for (const title of [null, "Revenue"]) {
          for (const width of [240, 360, 480, 600]) {
            for (const aspect of [0.4, ASPECT, 1]) {
              const data = chartOf(kind, points, series, { title });
              const reach = drawnAt(data, box(width, aspect));
              const frame = frameFor(data, reach, aspect, width);
              const why = `${kind}/${String(series)} ${String(points)} ${String(title)} ${String(width)} ${String(aspect)}`;
              expect(drawnAt(data, frame).left, why).toBeCloseTo(reach.left, 0);
              expect(drawnAt(data, frame).top, why).toBeCloseTo(reach.top, 0);
              expect(drawnAt(data, frame).width, why).toBeCloseTo(
                reach.width,
                0,
              );
              expect(frame.width, why).toBeCloseTo(width, 0);
            }
          }
        }
      }
    }
  });

  it("keeps the corner of a chart whose data grew by a point", () => {
    for (const [kind, series] of KINDS) {
      const old = chartOf(kind, 6, series, { title: null });
      const grown = chartOf(kind, 7, series, { title: null });
      const reach = drawnAt(old, box(360));
      const frame = frameFor(grown, reach, ASPECT, 360);
      expect(drawnAt(grown, frame).left, kind).toBeCloseTo(reach.left, 0);
      expect(drawnAt(grown, frame).top, kind).toBeCloseTo(reach.top, 0);
    }
  });

  it("does not drift over repeated rebuilds", () => {
    for (const [kind, series] of KINDS) {
      const data = chartOf(kind, 6, series, { title: null });
      const first = drawnAt(data, box(360));
      let reach = first;
      for (let update = 0; update < 4; update += 1) {
        reach = drawnAt(data, frameFor(data, reach, ASPECT, 360));
      }
      expectReach(reach, first);
    }
  });
});

describe("a pie whose legend rows set its width", () => {
  const data = chartOf("pie", 6, 1, { title: null });
  // Four legend items a row: 305 points wide at any box from 320 up to 400.
  const reach = drawnAt(data, box(360));

  it("takes the natural width when it draws that reach", () => {
    expect(reach.width).toBe(305);
    expect(frameFor(data, reach, ASPECT, 360).width).toBeCloseTo(360, 6);
    expect(frameFor(data, reach, ASPECT, 330).width).toBeCloseTo(330, 6);
  });

  it("falls to the nearest width that does when the natural one does not", () => {
    const narrow = frameFor(data, reach, ASPECT, 100);
    const wide = frameFor(data, reach, ASPECT, 1000);
    expect(narrow.width).toBeCloseTo(320, 1);
    expect(wide.width).toBeGreaterThan(399);
    expect(wide.width).toBeLessThan(400);
    expect(drawnAt(data, narrow).width).toBeCloseTo(305, 0);
    expect(drawnAt(data, wide).width).toBeCloseTo(305, 0);
    expect(drawnAt(data, wide).top).toBeCloseTo(reach.top, 0);
  });
});

describe("a chart with nothing to frame", () => {
  it("is the reach itself when the layout draws nothing", () => {
    // A kind with no shape version draws no primitives.
    const scatter = chartOf("scatter", 4, 2);
    const frame = frameFor(
      scatter,
      { left: 50, top: 60, width: 300 },
      0.5,
      300,
    );
    expect(frame).toEqual({ left: 50, top: 60, width: 300, height: 150 });
  });

  it("is the reach itself when the reach has no width", () => {
    const column = chartOf("column", 6, 2, { title: null });
    const frame = frameFor(column, { left: 50, top: 60, width: 0 }, 0.5, 300);
    expect(frame).toEqual({ left: 50, top: 60, width: 0, height: 0 });
  });
});
