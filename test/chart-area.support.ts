// Shared fixtures of the area layout suites (src/chart-shapes-area*.test.ts):
// a chart of seeded series, the filters that pick a kind of primitive out of
// a layout, and the plot and the value-to-slide maths a test works its
// expected numbers out with. Owns no assertions.

import {
  bands,
  valueY,
  type Primitive,
  type Rect,
  type Text,
  type ValueScale,
} from "../src/chart-shapes";
import type { Triangle } from "../src/chart-shapes-types";
import type { Box } from "../src/layout";
import type { ChartData, ChartSeries } from "../src/link/chart-model";

export const BOX: Box = { left: 10, top: 20, width: 480, height: 270 };
export const DARK = "#14213D";
export const LIGHT = "#F4A259";
export const INK = "#111111";
export const WHITE = "#FFFFFF";

export interface Seed {
  values: (number | null)[];
  color?: string;
  extra?: Partial<ChartSeries>;
}

function seriesOf(seed: Seed, index: number): ChartSeries {
  return {
    name: `s${String(index)}`,
    values: seed.values,
    labels: seed.values.map((value) => (value === null ? "" : String(value))),
    colors: seed.values.map(() => seed.color ?? LIGHT),
    ...seed.extra,
  };
}

export function areaOf(
  kind: "area" | "stackedArea",
  seeds: Seed[],
  extra: Partial<ChartData> = {},
): ChartData {
  const points = seeds[0]!.values.length;
  return {
    v: 1,
    kind,
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
export const triangles = (out: Primitive[]): Triangle[] =>
  out.filter((one): one is Triangle => one.kind === "triangle");
// Which end of its box a triangle's tip is at, and where the tip itself is.
export const tipEnd = (one: Triangle): "left" | "right" =>
  one.apex > 0.5 ? "right" : "left";
export const tipX = (one: Triangle): number =>
  one.box.left + one.apex * one.box.width;
export const texts = (out: Primitive[]): Text[] =>
  out.filter((one): one is Text => one.kind === "text");
export const named = (out: Primitive[], prefix: string): Primitive[] =>
  out.filter((one) => one.name.startsWith(prefix));
// The rectangles of the fills, not the legend's swatches.
export const fillRects = (out: Primitive[]): Rect[] =>
  rects(out).filter((one) => !one.name.startsWith("legend"));
// The fills in drawing order: rectangles and triangles, not the legend.
export const fills = (out: Primitive[]): (Rect | Triangle)[] =>
  out.filter(
    (one): one is Rect | Triangle =>
      (one.kind === "rect" || one.kind === "triangle") &&
      !one.name.startsWith("legend"),
  );

// The plot of a chart in BOX and where a value or a point lands in it, for a
// scale the test states itself (the layout works its own out from the data).
export function frame(data: ChartData, scale: ValueScale) {
  const { plot } = bands(data, BOX);
  const slot = plot.width / data.categories.length;
  return {
    plot,
    slot,
    x: (i: number) => plot.left + (i + 0.5) * slot,
    y: (value: number) => valueY(value, scale, plot),
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
