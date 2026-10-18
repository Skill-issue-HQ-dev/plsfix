// Shared fixtures of the combo layout suites (src/chart-shapes-combo*.test.ts):
// a combo of seeded series, the filters that pick a kind of primitive out of a
// layout, and the plot and the value-to-slide maths a test works its expected
// numbers out with. Owns no assertions.

import {
  bands,
  type Ellipse,
  type Line,
  type Primitive,
  type Rect,
  type Text,
} from "../src/chart-shapes";
import { COMBO_HEADROOM } from "../src/chart-shapes-combo";
import type { Box } from "../src/layout";
import type { ChartData } from "../src/link/chart-model";
import type {
  ChartSeries,
  SeriesAxis,
  SeriesType,
} from "../src/link/chart-series";

export const BOX: Box = { left: 10, top: 20, width: 480, height: 270 };
export const INK = "#111111";
export const BRASS = "#B8902D";
export const NAVY = "#14213D";
export const SKY = "#98C1D9";

export interface Seed {
  type: SeriesType;
  values: (number | null)[];
  axis?: SeriesAxis;
  color?: string;
  // The text of each label; the value as written when not given.
  labels?: string[];
  extra?: Partial<ChartSeries>;
}

function seriesOf(seed: Seed, index: number): ChartSeries {
  const color = seed.color ?? [BRASS, NAVY, SKY][index % 3]!;
  return {
    name: `s${String(index)}`,
    type: seed.type,
    values: seed.values,
    labels:
      seed.labels ??
      seed.values.map((value) => (value === null ? "" : String(value))),
    colors: seed.values.map(() => color),
    ...(seed.axis === undefined ? {} : { axis: seed.axis }),
    ...seed.extra,
  };
}

export function comboOf(
  seeds: Seed[],
  extra: Partial<ChartData> = {},
): ChartData {
  const points = seeds[0]!.values.length;
  return {
    v: 1,
    kind: "combo",
    title: null,
    categories: Array.from({ length: points }, (_one, i) => `C${String(i)}`),
    series: seeds.map(seriesOf),
    font: "Arial",
    ink: INK,
    titleColor: "#000000",
    ...extra,
  };
}

export const rects = (out: Primitive[]): Rect[] =>
  out.filter((one): one is Rect => one.kind === "rect");
export const bars = (out: Primitive[]): Rect[] =>
  rects(out).filter((one) => one.name.startsWith("bar "));
export const markers = (out: Primitive[]): Ellipse[] =>
  out.filter((one): one is Ellipse => one.kind === "ellipse");
export const segments = (out: Primitive[]): Line[] =>
  out.filter(
    (one): one is Line => one.kind === "line" && one.name.startsWith("line "),
  );
export const texts = (out: Primitive[]): Text[] =>
  out.filter((one): one is Text => one.kind === "text");
export const labels = (out: Primitive[]): Text[] =>
  texts(out).filter((one) => one.name.startsWith("label "));
export const named = (out: Primitive[], name: string): Primitive =>
  out.find((one) => one.name === name)!;

export const centreOf = (one: Primitive): { x: number; y: number } => ({
  x: one.box.left + one.box.width / 2,
  y: one.box.top + one.box.height / 2,
});

// The plot of a chart in BOX, the headroom the layout keeps over its tallest
// value for a label, and where a value lands on a scale the test states.
export function frame(data: ChartData) {
  const { plot } = bands(data, BOX);
  const room = Math.min(COMBO_HEADROOM, plot.height / 3);
  return {
    plot,
    room,
    // The slide's y for `value` on an axis that runs from `min` to `max`: the
    // top of the axis is `room` under the plot's top, the bottom is the plot's.
    y: (value: number, max: number, min = 0): number =>
      plot.top +
      room +
      ((max - value) / (max - min || 1)) * (plot.height - room),
    slot: plot.width / data.categories.length,
    x: (i: number): number =>
      plot.left + (i + 0.5) * (plot.width / data.categories.length),
  };
}

export function insideBox(one: Primitive, box: Box, slack = 0.05): boolean {
  const { left, top, width, height } = one.box;
  return (
    width >= 0 &&
    height >= 0 &&
    left >= box.left - slack &&
    top >= box.top - slack &&
    left + width <= box.left + box.width + slack &&
    top + height <= box.top + box.height + slack
  );
}
