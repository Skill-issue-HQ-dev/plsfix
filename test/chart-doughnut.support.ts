// Shared fixtures of the doughnut layout suites (src/chart-shapes-doughnut*.test.ts):
// a doughnut of seeded rings, the filters that pick a kind of primitive out of
// a layout, and the circle maths a test works its expected numbers out with.
// Owns no assertions.

import { bands, type Primitive, type Text } from "../src/chart-shapes";
import type { BlockArc, Donut } from "../src/chart-shapes-types";
import type { Box } from "../src/layout";
import type { ChartData } from "../src/link/chart-model";

export const BOX: Box = { left: 10, top: 20, width: 480, height: 270 };
export const INK = "#111111";
export const DARK = "#14213D";
export const LIGHT = "#F4A259";
export const WHITE = "#FFFFFF";
export const COLORS = [DARK, LIGHT, "#2EC4B6", "#B27E54", "#8AB17D", "#E63946"];

// One ring's series: a value per slice, its text, a colour per slice.
function ringOf(values: (number | null)[]) {
  return {
    name: "ring",
    values,
    labels: values.map((value) => (value === null ? "" : String(value))),
    colors: values.map((_value, i) => COLORS[i % COLORS.length]!),
  };
}

export function doughnut(
  rings: (number | null)[][],
  extra: Partial<ChartData> = {},
): ChartData {
  const points = rings[0]!.length;
  return {
    v: 1,
    kind: "doughnut",
    title: null,
    categories: Array.from({ length: points }, (_one, i) => `Slice ${i}`),
    series: rings.map((values, j) => ({
      ...ringOf(values),
      name: `Ring ${String(j)}`,
    })),
    font: "Arial",
    ink: INK,
    titleColor: "#000000",
    ...extra,
  };
}

export const arcs = (out: Primitive[]): BlockArc[] =>
  out.filter((one): one is BlockArc => one.kind === "blockArc");
export const donuts = (out: Primitive[]): Donut[] =>
  out.filter((one): one is Donut => one.kind === "donut");
export const labels = (out: Primitive[]): Text[] =>
  out.filter(
    (one): one is Text => one.kind === "text" && one.name.startsWith("label "),
  );

// How far an arc sweeps clockwise from its start to its end, however the
// host's fold into (-180, 180] split it across the 3 o'clock line.
export const sweep = (arc: BlockArc): number =>
  (((arc.end - arc.start) % 360) + 360) % 360;

// The doughnut's circle in BOX: its centre and outer radius.
export function circle(data: ChartData) {
  const { plot } = bands(data, BOX);
  return {
    cx: plot.left + plot.width / 2,
    cy: plot.top + plot.height / 2,
    radius: Math.min(plot.width, plot.height) / 2,
  };
}

// An arc's inner and outer radius as the host will draw it: the box is the
// ring's square, the thickness a fraction of its side.
export const radii = (arc: BlockArc | Donut) => {
  const outer = arc.box.width / 2;
  return { outer, inner: outer - arc.thickness * arc.box.width };
};
