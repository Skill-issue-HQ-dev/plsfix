// Hunt pass, C5: the two kinds with maths of their own, as properties over
// random charts at the caps (up to 12 series, 24 slices a ring, 120 points).
// A doughnut: every ring's slices close it exactly (sweeps add up to 360, a lone
// slice is a Donut, none for none), every angle is one the host reads back, no
// thickness fills a ring solid, and the rings lie edge to edge from the hole to
// the doughnut's rim. A combo: every shape is named for a series of the role it
// has (bars for columns, segments and markers for lines), and where the plot has
// room for a label, no label touches the title. The generic invariants (inside
// the box, a name of its own, never empty) cover both kinds in
// chart-shapes.properties.hunt2.test.ts.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  bands,
  layoutChart,
  MIN_SIZE,
  TITLE_BAND,
} from "../../src/chart-shapes";
import { COMBO_HEADROOM } from "../../src/chart-shapes-combo";
import type { BlockArc, Donut } from "../../src/chart-shapes-types";
import type { Primitive } from "../../src/chart-shapes";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  normalizeAngle,
  type ChartData,
} from "../../src/link/chart-model";
import type { ChartSeries, SeriesType } from "../../src/link/chart-series";

const SEED = 20261005;
const NUM_RUNS = 300;
const EPS = 0.05;

const cell = fc.option(fc.integer({ min: -3, max: 1000 }), {
  nil: null,
  freq: 5,
});

function seriesOf(
  values: (number | null)[],
  j: number,
  extra: Partial<ChartSeries> = {},
): ChartSeries {
  return {
    name: `S${String(j)}`,
    values,
    labels: values.map((one) => (one === null ? "" : String(one))),
    colors: values.map((_one, i) => (i % 2 === 0 ? "#14213D" : "#F4A259")),
    ...extra,
  };
}

const boxArb = fc.record({
  left: fc.integer({ min: 0, max: 200 }),
  top: fc.integer({ min: 0, max: 200 }),
  width: fc.integer({ min: MIN_SIZE.width, max: 900 }),
  height: fc.integer({ min: MIN_SIZE.height, max: 500 }),
});

const FRAME = {
  v: 1 as const,
  title: "T",
  font: "Arial",
  ink: "#282623",
  titleColor: "#14213D",
};

// A doughnut of 1 to 12 rings over 1 to 24 slices, any hole or none.
const doughnutArb = fc
  .tuple(
    fc.integer({ min: 1, max: 24 }),
    fc.integer({ min: 1, max: CHART_MAX_SERIES }),
    fc.option(fc.integer({ min: 10, max: 90 }), { nil: undefined }),
  )
  .chain(([points, rings, hole]) =>
    fc
      .array(fc.array(cell, { minLength: points, maxLength: points }), {
        minLength: rings,
        maxLength: rings,
      })
      .map((all): ChartData => ({
        ...FRAME,
        kind: "doughnut",
        categories: Array.from({ length: points }, (_one, i) => `C${i}`),
        series: all.map((values, j) => seriesOf(values, j)),
        ...(hole === undefined ? {} : { hole }),
      })),
  );

const radiiOf = (one: BlockArc | Donut) => {
  const outer = one.box.width / 2;
  return { outer, inner: outer - one.thickness * one.box.width };
};

const ringOf = (out: Primitive[], j: number): (BlockArc | Donut)[] =>
  out.filter(
    (one): one is BlockArc | Donut =>
      (one.kind === "blockArc" || one.kind === "donut") &&
      one.name.startsWith(`slice ${String(j)}.`),
  );

describe("layoutChart: a doughnut's rings", () => {
  it("closes every ring exactly: sweeps add up to 360, a lone slice is a Donut, no positive slice draws none", () => {
    fc.assert(
      fc.property(doughnutArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        data.series.forEach((series, j) => {
          const positive = series.values.filter(
            (one) => one !== null && one > 0,
          );
          const ring = ringOf(out, j);
          const arcs = ring.filter(
            (one): one is BlockArc => one.kind === "blockArc",
          );
          const whole = ring.filter(
            (one): one is Donut => one.kind === "donut",
          );
          if (positive.length === 0) {
            expect(ring).toEqual([]);
          } else if (positive.length === 1) {
            expect(arcs).toEqual([]);
            expect(whole).toHaveLength(1);
          } else {
            expect(arcs).toHaveLength(positive.length);
            expect(whole).toEqual([]);
            const total = arcs.reduce(
              (sum, one) => sum + ((((one.end - one.start) % 360) + 360) % 360),
              0,
            );
            expect(Math.abs(total - 360)).toBeLessThan(0.01);
          }
        });
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("writes only angles the host reads back, two different ones for a slice, and a thickness under a half", () => {
    fc.assert(
      fc.property(doughnutArb, boxArb, (data, box) => {
        for (const one of layoutChart(data, box)) {
          if (one.kind === "blockArc") {
            for (const angle of [one.start, one.end]) {
              expect(angle).toBe(normalizeAngle(angle));
            }
            expect(one.start).not.toBe(one.end);
          }
          if (one.kind === "blockArc" || one.kind === "donut") {
            expect(one.thickness).toBeGreaterThanOrEqual(0);
            expect(one.thickness).toBeLessThan(0.5);
            expect(one.box.width).toBeCloseTo(one.box.height, 6);
          }
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("lays the rings edge to edge from the hole to the rim, the first series innermost", () => {
    fc.assert(
      fc.property(doughnutArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        const { plot } = bands(data, box);
        const rim = Math.min(plot.width, plot.height) / 2;
        const hole = (data.hole ?? 75) / 100;
        const drawn = data.series
          .map((_one, j) => ringOf(out, j)[0])
          .map((one, j) => ({ j, one }))
          .filter(
            (entry): entry is { j: number; one: BlockArc | Donut } =>
              entry.one !== undefined,
          );
        for (const { j, one } of drawn) {
          const width = (rim * (1 - hole)) / data.series.length;
          const { inner, outer } = radiiOf(one);
          expect(inner).toBeCloseTo(rim * hole + j * width, 6);
          expect(outer).toBeCloseTo(rim * hole + (j + 1) * width, 6);
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});

// A drawable combo: columns (all clustered, or all stacked) and lines, a line on
// either axis, at least two series.
const comboArb = fc
  .tuple(
    fc.integer({ min: 1, max: CHART_MAX_POINTS }),
    fc.integer({ min: 2, max: CHART_MAX_SERIES }),
    fc.boolean(),
  )
  .chain(([points, count, stacked]) =>
    fc
      .array(fc.array(cell, { minLength: points, maxLength: points }), {
        minLength: count,
        maxLength: count,
      })
      .map((all): ChartData => {
        const types: SeriesType[] = all.map((_one, j) =>
          j % 3 === 1 || j % 3 === 2
            ? "line"
            : stacked
              ? "stackedColumn"
              : "column",
        );
        return {
          ...FRAME,
          kind: "combo",
          categories: Array.from({ length: points }, (_one, i) => `C${i}`),
          series: all.map((values, j) =>
            seriesOf(values, j, {
              type: types[j]!,
              ...(j % 3 === 2 ? { axis: "secondary" as const } : {}),
            }),
          ),
        };
      }),
  );

describe("layoutChart: a combo's shapes", () => {
  it("names bars for the column series and segments, markers and labels' series only for what they belong to", () => {
    fc.assert(
      fc.property(comboArb, boxArb, (data, box) => {
        for (const one of layoutChart(data, box)) {
          const match = /^(bar|line|marker|label) (\d+)\./.exec(one.name);
          if (match === null) continue;
          const type = data.series[Number(match[2])]?.type;
          expect(type, one.name).toBeDefined();
          if (match[1] === "bar") {
            expect(["column", "stackedColumn"], one.name).toContain(type);
          } else if (match[1] !== "label") {
            expect(type, one.name).toBe("line");
          }
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("keeps every label off the title wherever the plot has room for one", () => {
    fc.assert(
      fc.property(comboArb, boxArb, (data, box) => {
        const { plot } = bands(data, box);
        fc.pre(plot.height >= 3 * COMBO_HEADROOM);
        for (const one of layoutChart(data, box)) {
          if (one.kind === "text" && one.name.startsWith("label ")) {
            expect(one.box.top, one.name).toBeGreaterThanOrEqual(
              box.top + TITLE_BAND - EPS,
            );
          }
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("draws a bar only where a column series has a value", () => {
    fc.assert(
      fc.property(comboArb, boxArb, (data, box) => {
        const bars = layoutChart(data, box).filter(
          (one) => one.kind === "rect" && one.name.startsWith("bar "),
        );
        const present = data.series
          .filter((one) => one.type !== "line")
          .reduce(
            (sum, one) => sum + one.values.filter((v) => v !== null).length,
            0,
          );
        expect(bars).toHaveLength(present);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});

// The slices cap is a payload's, not the layout's: a doughnut of exactly the
// caps still lays out.
describe("layoutChart: a doughnut at the caps", () => {
  it("draws twelve rings of 24 slices", () => {
    const values = Array.from({ length: CHART_MAX_SLICES }, (_one, i) => i + 1);
    const data: ChartData = {
      ...FRAME,
      kind: "doughnut",
      categories: values.map((one) => `C${String(one)}`),
      series: Array.from({ length: CHART_MAX_SERIES }, (_one, j) =>
        seriesOf(values, j),
      ),
    };
    const out = layoutChart(data, { left: 0, top: 0, width: 600, height: 400 });
    expect(out.filter((one) => one.kind === "blockArc")).toHaveLength(
      CHART_MAX_SERIES * CHART_MAX_SLICES,
    );
  });
});
