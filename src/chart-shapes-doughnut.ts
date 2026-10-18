// The doughnut kind: one BlockArc per slice on the square box of its ring, the
// slices cut clockwise from 12 o'clock like a pie's, several series as
// concentric rings with the first innermost (as Excel draws them) and a lone
// 100 % slice as a Donut. The hole is Excel's doughnutHoleSize, a share of the
// outer diameter. A slice's value sits at the middle of its ring in the ink that
// reads on its colour, and is left out where the slice is too thin to hold it;
// a zero, a negative and a blank slice draw nothing and take no angle. Pure: no
// Office.js. Invariant: every primitive lies inside the box it was given.

import {
  LABEL_HEIGHT,
  LABEL_PAD,
  LABEL_SIZE,
  MIN_SEGMENT,
  bands,
  darkFill,
  textWidth,
  type Primitive,
} from "./chart-shapes";
import { boxAt, label, legend, titleText } from "./chart-shapes-parts";
import type { BlockArc, Donut } from "./chart-shapes-types";
import type { Box } from "./layout";
import {
  holeFraction,
  normalizeAngle,
  type ChartData,
} from "./link/chart-model";
import {
  positiveTotal,
  showsLabels,
  type ChartSeries,
} from "./link/chart-series";

// 12 o'clock, in the host's angles: degrees clockwise from 3 o'clock.
const START_DEGREES = -90;
const DEGREES_PER_TURN = 360;
const DEGREES_TO_RADIANS = Math.PI / 180;
const DOUBLE = 2;
const WHITE = "#FFFFFF";
// Slack for a corner of a label that rounding put a hair outside its slice.
const EDGE = 1e-6;
// The thinnest slice that is an angle: a billionth of a degree, a thousandth of
// a micron at the rim of the largest doughnut, and a thousand times what a
// double can tell from its neighbour at 360, so a start never equals its end.
const MIN_SPAN = 1e-9;

// The doughnut's circle: its centre and outer radius, and the plot it sits in.
interface Disc {
  cx: number;
  cy: number;
  radius: number;
  plot: Box;
}

// One ring between two radii of the disc.
interface Ring {
  inner: number;
  outer: number;
}

// One drawn slice: the category it is, and where it starts and ends (degrees
// clockwise from 3 o'clock, not folded). A slice ends exactly where the next
// one starts, and the last at the full turn.
interface Slice {
  index: number;
  from: number;
  to: number;
}

function discOf(plot: Box): Disc {
  return {
    cx: plot.left + plot.width / DOUBLE,
    cy: plot.top + plot.height / DOUBLE,
    radius: Math.min(plot.width, plot.height) / DOUBLE,
    plot,
  };
}

// The rings from the hole out: the first series innermost, the last at the
// doughnut's edge, each as wide as the others (Excel's rings are).
function ringsOf(count: number, radius: number, hole: number): Ring[] {
  const width = (radius * (1 - hole)) / count;
  return Array.from({ length: count }, (_one, k) => {
    const inner = radius * hole + k * width;
    return { inner, outer: inner + width };
  });
}

// The slices of one series from 12 o'clock, clockwise, each as wide as its
// share of the positive total. A zero, a negative and a blank cell are no
// slice and take no angle; nor is a slice thinner than MIN_SPAN, which is no
// angle at all (the host would read a start equal to its end as a full sweep).
// Each slice's end is the next one's start (the same number, not a sum rounded
// twice) and the last runs to the full turn, so the ring closes with no seam
// and no overlap.
function slicesOf(series: ChartSeries): Slice[] {
  const total = positiveTotal(series.values);
  if (total <= 0) return [];
  const drawn = series.values.flatMap((value, index) =>
    value === null || value <= 0 ? [] : [{ index, value }],
  );
  let edge = START_DEGREES;
  return drawn
    .map(({ index, value }, k) => {
      const from = edge;
      edge =
        k === drawn.length - 1
          ? START_DEGREES + DEGREES_PER_TURN
          : from + (value / total) * DEGREES_PER_TURN;
      return { index, from, to: edge };
    })
    .filter((slice) => slice.to - slice.from >= MIN_SPAN);
}

// The shape of one slice on its ring: a Donut for the lone slice, else a
// BlockArc.
function shapeOf(
  slice: Slice,
  lone: boolean,
  ring: Ring,
  disc: Disc,
  series: ChartSeries,
  seriesIndex: number,
): BlockArc | Donut {
  const { outer, inner } = ring;
  const box = boxAt(
    disc.cx - outer,
    disc.cy - outer,
    outer * DOUBLE,
    outer * DOUBLE,
  );
  const thickness = outer > 0 ? (outer - inner) / (outer * DOUBLE) : 0;
  const color = series.colors[slice.index]!;
  const name = `slice ${String(seriesIndex)}.${String(slice.index)}`;
  if (lone) return { kind: "donut", box, thickness, color, name };
  const start = normalizeAngle(slice.from);
  const end = normalizeAngle(slice.to);
  return { kind: "blockArc", box, start, end, thickness, color, name };
}

// Whether the text, centred at (x, y), lies whole inside the slice: each of its
// four corners between the ring's two radii and inside its sweep.
function holds(
  x: number,
  y: number,
  half: { width: number; height: number },
  slice: Slice,
  ring: Ring,
  disc: Disc,
): boolean {
  return [-1, 1].every((sx) =>
    [-1, 1].every((sy) => {
      const dx = x + sx * half.width - disc.cx;
      const dy = y + sy * half.height - disc.cy;
      const r = Math.hypot(dx, dy);
      if (r < ring.inner - EDGE || r > ring.outer + EDGE) return false;
      const angle = (Math.atan2(dy, dx) * DEGREES_PER_TURN) / (2 * Math.PI);
      const along =
        (((angle - slice.from) % DEGREES_PER_TURN) + DEGREES_PER_TURN) %
        DEGREES_PER_TURN;
      return along <= slice.to - slice.from + EDGE;
    }),
  );
}

function clamp(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(value, high));
}

// A slice's value at the middle of its ring, on the slice's middle angle, in
// white over a dark colour and the chart's ink over a light one; none when the
// series shows no labels, the cell has no text, or the text does not fit.
function labelOf(
  slice: Slice,
  ring: Ring,
  disc: Disc,
  series: ChartSeries,
  seriesIndex: number,
  ink: string,
): Primitive | null {
  const text = series.labels[slice.index] ?? "";
  if (!showsLabels(series) || text === "") return null;
  const middle = ((slice.from + slice.to) / DOUBLE) * DEGREES_TO_RADIANS;
  const radius = (ring.inner + ring.outer) / DOUBLE;
  const x = disc.cx + radius * Math.cos(middle);
  const y = disc.cy + radius * Math.sin(middle);
  const glyphs = textWidth(text, LABEL_SIZE) - LABEL_PAD;
  const half = { width: glyphs / DOUBLE, height: MIN_SEGMENT / DOUBLE };
  if (!holds(x, y, half, slice, ring, disc)) return null;
  const { plot } = disc;
  const width = Math.min(textWidth(text, LABEL_SIZE), plot.width);
  const left = clamp(
    x - width / DOUBLE,
    plot.left,
    plot.left + plot.width - width,
  );
  const top = clamp(
    y - LABEL_HEIGHT / DOUBLE,
    plot.top,
    plot.top + plot.height - LABEL_HEIGHT,
  );
  const fill = series.colors[slice.index]!;
  const color = darkFill(fill) ? WHITE : ink;
  const name = `label ${String(seriesIndex)}.${String(slice.index)}`;
  return label(boxAt(left, top, width, LABEL_HEIGHT), text, color, "c", name);
}

// A doughnut as primitives: the title, every ring's slices (the first series
// innermost), their values over them, and the legend of the categories.
export function layoutDoughnut(data: ChartData, box: Box): Primitive[] {
  const { title, plot, legendBand } = bands(data, box);
  const disc = discOf(plot);
  const rings = ringsOf(data.series.length, disc.radius, holeFraction(data));
  const shapes: Primitive[] = [];
  const values: Primitive[] = [];
  data.series.forEach((series, j) => {
    const ring = rings[j]!;
    const slices = slicesOf(series);
    for (const slice of slices) {
      shapes.push(shapeOf(slice, slices.length === 1, ring, disc, series, j));
      const text = labelOf(slice, ring, disc, series, j, data.ink);
      if (text !== null) values.push(text);
    }
  });
  const out: Primitive[] = [];
  if (data.title !== null) {
    out.push(titleText(title, data.title, data.titleColor));
  }
  out.push(...shapes, ...values);
  if (legendBand) out.push(...legend(data, legendBand));
  return out;
}
