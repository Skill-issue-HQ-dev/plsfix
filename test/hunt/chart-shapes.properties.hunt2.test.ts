// Hunt pass 2, target 3: layoutChart (src/chart-shapes.ts, the one entry
// point every chart-shapes*.ts file feeds, through the registry in
// src/chart-kinds.ts) as a property across every drawable chart kind, random
// series at the new caps (up to 120 points and 12 series), any sign, zeros,
// gaps (blank cells), 100 % stacks, totals, reversed category axes, per-series
// flags, and huge or tiny magnitudes. Plain-word invariant, the one every
// existing chart-shapes suite already states in its own header comment: every
// primitive lies inside the box it was given, with a non-negative width and
// height (PowerPoint throws InvalidArgument on a negative side, on both an
// add and a later write), and every primitive has a name of its own.
// Pass 1's edge tests (ppt.chart-shapes.edges.hunt) hand-picked single-series
// scenarios per kind; this generates thousands of them and lets fast-check
// shrink any counterexample to its smallest form.

import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { displayOrder } from "../../src/chart-kinds";
import { layoutChart, MIN_SIZE, type Primitive } from "../../src/chart-shapes";
import type { Box } from "../../src/layout";
import {
  CHART_MAX_POINTS,
  CHART_MAX_SERIES,
  CHART_MAX_SLICES,
  type ChartData,
  type ChartKind,
  type ChartSeries,
} from "../../src/link/chart-model";

const SEED = 20260927;
const NUM_RUNS = 400;
const EPS = 0.05;

// Every kind that has a layout; the one that has none (a scatter) draws nothing
// (below). Since C5 that is every kind but one: the doughnut (up to 12 rings of
// 24 slices) and the combo (columns clustered or stacked, lines on either axis)
// are generated here too, their own invariants in
// chart-shapes.rings-combos.hunt2.test.ts.
const DRAWN: ChartKind[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "waterfall",
  "pie",
  "line",
  "doughnut",
  "area",
  "stackedArea",
  "combo",
];
const NOT_YET: ChartKind[] = ["scatter"];

// Any sign, zero, and both huge and tiny finite magnitudes - what a real
// modeller's sheet can hold and Excel will still export as a number.
const extremeValue = fc.oneof(
  fc.integer({ min: -1_000_000, max: 1_000_000 }),
  fc.constant(0),
  fc.constant(-0),
  fc.double({ min: 1, max: 1e300, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: -1e300, max: -1, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: 1e-300, max: 1e-6, noNaN: true, noDefaultInfinity: true }),
  fc.double({ min: -1e-6, max: -1e-300, noNaN: true, noDefaultInfinity: true }),
);

// A blank cell, one in six: a gap is as ordinary as a number.
const cell = fc.option(extremeValue, { nil: null, freq: 6 });

const HEX_COLORS = [
  "#2EC4B6",
  "#B27E54",
  "#14213D",
  "#E63946",
  "#8AB17D",
  "#F4A259",
];

function seriesArb(points: number): fc.Arbitrary<ChartSeries> {
  return fc
    .tuple(
      fc.string({ minLength: 1, maxLength: 10 }),
      fc.array(cell, { minLength: points, maxLength: points }),
      fc.array(fc.constantFrom(...HEX_COLORS), {
        minLength: points,
        maxLength: points,
      }),
      fc.boolean(),
      fc.boolean(),
    )
    .map(([name, values, colors, dataLabels, markers]): ChartSeries => ({
      name,
      values,
      labels: values.map((v) => (v === null ? "" : String(v))),
      colors,
      ...(dataLabels ? {} : { dataLabels: false }),
      ...(markers ? {} : { markers: false }),
    }));
}

// kind, category count and series count first, so every series generated
// afterwards shares exactly the same point count (layoutChart, like real
// Excel exports, assumes every series in one chart has the same length).
//
// OUT-OF-SLICE, not worked around here by accident: chartmath.ts's
// bridgeSeries (src/chartmath.ts:26-34, not owned by this slice) throws
// below 3 points ("A bridge needs an opening total, a delta and a closing
// total."), but link/chart-model.ts's CHART_MIN_POINTS (line 40, also not
// owned) admits a "waterfall" at the blanket minimum of 2 like every other
// kind - a real mismatch a corrupted or future payload could hit, reported
// under OUT-OF-SLICE. Generating waterfall at >= 3 here tests this file's
// OWN invariant over the inputs valueScale (chart-shapes.ts:178-179) can
// actually be called with, rather than re-failing on the other module's gap
// on every run.
const pointsFor = (kind: ChartKind): fc.Arbitrary<number> =>
  fc.integer({
    min: kind === "waterfall" ? 3 : 1,
    max:
      kind === "pie" || kind === "doughnut"
        ? CHART_MAX_SLICES
        : CHART_MAX_POINTS,
  });

// A pie and a waterfall are one series, a combo at least two (columns and
// lines); everything else up to twelve, a doughnut's twelve being rings.
const seriesCountFor = (kind: ChartKind): fc.Arbitrary<number> =>
  kind === "pie" || kind === "waterfall"
    ? fc.constant(1)
    : fc.integer({ min: kind === "combo" ? 2 : 1, max: CHART_MAX_SERIES });

// A drawable combo by place: the even series are columns (all clustered or all
// stacked, by the point count's parity), the odd ones lines, every second line
// on the secondary axis.
function asCombo(series: ChartSeries[], points: number): ChartSeries[] {
  const stacked = points % 2 === 0;
  return series.map((one, j): ChartSeries => {
    if (j % 2 === 0) {
      return { ...one, type: stacked ? "stackedColumn" : "column" };
    }
    return {
      ...one,
      type: "line",
      ...(j % 4 === 1 ? { axis: "secondary" as const } : {}),
    };
  });
}

// The optional chart-level fields a newer pane adds, on the kinds that take
// them: a 100 % stack, a total label per category, the category axis' order
// (absent, true or an explicit false), a tornado's overlap.
function extrasFor(
  kind: ChartKind,
  points: number,
): fc.Arbitrary<Partial<ChartData>> {
  const stacked = kind === "stackedColumn" || kind === "stackedBar";
  const stackedArea = kind === "stackedArea";
  return fc
    .tuple(
      fc.boolean(),
      fc.constantFrom<boolean | undefined>(undefined, true, false),
      fc.boolean(),
      fc.boolean(),
      fc.array(fc.string({ maxLength: 8 }), {
        minLength: points,
        maxLength: points,
      }),
    )
    .map(([percent, reversed, overlap, totals, labels]) => ({
      ...((stacked || stackedArea) && percent
        ? { percent: true as const }
        : {}),
      ...(reversed === undefined ? {} : { categoriesReversed: reversed }),
      ...(kind === "bar" && overlap ? { overlap: true as const } : {}),
      ...(stacked && totals ? { totals: labels } : {}),
      // A doughnut's hole, 10 to 90 by the point count, or left unsaid.
      ...(kind === "doughnut" && points % 3 !== 0
        ? { hole: 10 + ((points * 7) % 81) }
        : {}),
    }));
}

function chartDataArb(kinds: ChartKind[]): fc.Arbitrary<ChartData> {
  return fc
    .tuple(
      fc.constantFrom(...kinds),
      fc.option(fc.string({ minLength: 1, maxLength: 20 }), { nil: null }),
    )
    .chain(([kind, title]) =>
      pointsFor(kind).chain((points) =>
        fc
          .tuple(
            seriesCountFor(kind).chain((count) =>
              fc.array(seriesArb(points), {
                minLength: count,
                maxLength: count,
              }),
            ),
            extrasFor(kind, points),
          )
          .map(([series, extras]): ChartData => ({
            v: 1,
            kind,
            title,
            categories: Array.from(
              { length: points },
              (_, i) => `Cat ${String(i)}`,
            ),
            series: kind === "combo" ? asCombo(series, points) : series,
            font: "Aptos Narrow",
            ink: "#282623",
            titleColor: "#14213D",
            ...extras,
          })),
      ),
    );
}

const drawnArb = chartDataArb(DRAWN);

// Comfortably at or above MIN_SIZE on both sides, the one precondition every
// real caller (charts.ts belowMinimum) already guarantees before a chart is
// ever laid out at all.
const boxArb = fc.record({
  left: fc.integer({ min: 0, max: 200 }),
  top: fc.integer({ min: 0, max: 200 }),
  width: fc.integer({ min: MIN_SIZE.width, max: 900 }),
  height: fc.integer({ min: MIN_SIZE.height, max: 500 }),
});

function offenders(primitives: Primitive[], box: Box): Primitive[] {
  return primitives.filter((p) => {
    const { box: pb } = p;
    return (
      !Number.isFinite(pb.left) ||
      !Number.isFinite(pb.top) ||
      !Number.isFinite(pb.width) ||
      !Number.isFinite(pb.height) ||
      pb.width < -EPS ||
      pb.height < -EPS ||
      pb.left < box.left - EPS ||
      pb.top < box.top - EPS ||
      pb.left + pb.width > box.left + box.width + EPS ||
      pb.top + pb.height > box.top + box.height + EPS
    );
  });
}

describe("layoutChart: every primitive stays inside its box, non-negative sides", () => {
  it("holds across all nine drawn kinds, up to 120 points and 12 series, gaps, any sign, 100 % stacks, totals, reversed axes, huge and tiny values", () => {
    fc.assert(
      fc.property(drawnArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        expect(offenders(out, box)).toEqual([]);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("draws at least one primitive whatever the kind or values (never an empty chart)", () => {
    fc.assert(
      fc.property(drawnArb, boxArb, (data, box) => {
        const out = layoutChart(data, box);
        expect(out.length).toBeGreaterThan(0);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("gives every primitive a name of its own", () => {
    fc.assert(
      fc.property(drawnArb, boxArb, (data, box) => {
        const names = layoutChart(data, box).map((one) => one.name);
        expect(new Set(names).size).toBe(names.length);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});

// A blank cell is nothing: the columns, bars and stacked segments drawn are
// exactly the cells that hold a number.
describe("layoutChart: a blank cell draws no bar", () => {
  const barKinds = chartDataArb([
    "column",
    "bar",
    "stackedColumn",
    "stackedBar",
  ]);

  it("draws one bar per cell that holds a value, and none for a blank", () => {
    fc.assert(
      fc.property(barKinds, boxArb, (data, box) => {
        const bars = layoutChart(data, box).filter(
          (one) => one.kind === "rect" && one.name.startsWith("bar "),
        );
        const present = data.series.reduce(
          (sum, one) => sum + one.values.filter((v) => v !== null).length,
          0,
        );
        expect(bars).toHaveLength(present);
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});

// The areas plot a blank cell as a zero, like Excel, so a fill never breaks,
// but a value label belongs to a cell that holds a value. A stack is painted
// highest series first, a plain area first series first.
describe("layoutChart: the area kinds", () => {
  const areaKinds = chartDataArb(["area", "stackedArea"]);

  function stretch(name: string): { series: number; at: number } | null {
    const match = /^(?:area|slope) (\d+)\.(\d+)/.exec(name);
    return match ? { series: Number(match[1]), at: Number(match[2]) } : null;
  }

  it("writes a value label only at a cell that holds a value", () => {
    fc.assert(
      fc.property(areaKinds, boxArb, (data, box) => {
        const shown = displayOrder(data);
        for (const one of layoutChart(data, box)) {
          const match = /^label (\d+)\.(\d+)$/.exec(one.name);
          if (match === null) continue;
          const values = shown.series[Number(match[1])]!.values;
          expect(values[Number(match[2])]).not.toBeNull();
        }
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });

  it("paints the series in the order the kind asks for", () => {
    fc.assert(
      fc.property(areaKinds, boxArb, (data, box) => {
        const order = layoutChart(data, box).flatMap((one) => {
          const found = stretch(one.name);
          return found === null ? [] : [found.series];
        });
        const sorted = [...order].sort((a, b) => a - b);
        expect(order).toEqual(data.kind === "area" ? sorted : sorted.reverse());
      }),
      { seed: SEED, numRuns: NUM_RUNS },
    );
  });
});

// The kind the reader already writes but no slice has laid out yet (a scatter)
// has no layout: the registry answers null and layoutChart an empty list, which
// declineReason turns into the "as a picture" sentence.
describe("layoutChart: kinds with no layout yet", () => {
  it("draw nothing, for any payload the validator would take", () => {
    const series = (name: string, extra: Partial<ChartSeries> = {}) => ({
      name,
      values: [1, 2, 3],
      labels: ["1", "2", "3"],
      colors: HEX_COLORS.slice(0, 3),
      ...extra,
    });
    const box = { left: 0, top: 0, width: 480, height: 270 };
    for (const kind of NOT_YET) {
      const data: ChartData = {
        v: 1,
        kind,
        title: null,
        categories: ["a", "b", "c"],
        series: [
          series("s", { x: [1, 2, 3], type: "column" }),
          series("t", { x: [1, 2, 3], type: "line" }),
        ],
        font: "Arial",
        ink: "#333333",
        titleColor: "#14213D",
      };
      expect(layoutChart(data, box), kind).toEqual([]);
    }
  });
});
