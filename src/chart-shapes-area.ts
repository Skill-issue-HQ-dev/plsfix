// The area kinds, area and stacked area (also 100 %), as primitives: a
// rectangle and a triangle per stretch (chart-shapes-area-pieces.ts), a blank
// cell plotted as a zero, a stack in painter's order. Pure: no Office.js.
// Invariant: every primitive lies inside the box it was given.

import {
  LABEL_HEIGHT,
  LABEL_SIZE,
  MIN_SEGMENT,
  baselineLine,
  bands,
  textWidth,
  valueY,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import { inkOver } from "./chart-shapes-area-hit";
import { AREA_SEAM, slopePieces, type Piece } from "./chart-shapes-area-pieces";
import {
  axisLabels,
  pointAt,
  pointAxis,
  type PointAxis,
} from "./chart-shapes-axis";
import { spreadEnds, type End } from "./chart-shapes-ends";
import { boxAt, label, legend, titleText } from "./chart-shapes-parts";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";
import { showsLabels } from "./link/chart-series";

// One series as the area draws it: the value-space top of its fill at each
// point and the floor its band sits on there, which is the baseline for a plain
// area and the stack below for a stacked one. A blank cell is a zero, as Excel
// plots it (the fill dips to the floor, a stack keeps a band of no height), and
// `has` remembers which points hold a value, since only those carry a label.
interface Layer {
  index: number;
  tops: number[];
  floors: number[];
  has: boolean[];
}

// The plot the layers are drawn into, the maths from a value to the slide and
// where along the axis the points sit.
interface Frame {
  plot: Box;
  scale: ValueScale;
  axis: PointAxis;
}

function clamp(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(value, high));
}

function plainLayers(data: ChartData): Layer[] {
  return data.series.map((series, index) => ({
    index,
    tops: series.values.map((value) => value ?? 0),
    floors: series.values.map(() => 0),
    has: series.values.map((value) => value !== null),
  }));
}

// A 100 % stack divides each value by its category's absolute total first, so
// a category of mixed signs still adds up to one whole and a category of zeros
// stays flat. A value under zero is only approximated: each series is painted
// from the baseline up to its running total, which is the band between two
// totals only while no series is negative (the Excel export and the Insert
// note say so).
function stackedLayers(data: ChartData): Layer[] {
  const percent = data.percent === true;
  const wholes = data.categories.map((_name, i) =>
    data.series.reduce((sum, one) => sum + Math.abs(one.values[i] ?? 0), 0),
  );
  const running = data.categories.map(() => 0);
  return data.series.map((series, index) => {
    const floors = [...running];
    const tops = series.values.map((value, i) => {
      if (value === null) return running[i]!;
      const share = percent ? (wholes[i]! > 0 ? value / wholes[i]! : 0) : value;
      running[i] = running[i]! + share;
      return running[i]!;
    });
    const has = series.values.map((value) => value !== null);
    return { index, tops, floors, has };
  });
}

// From the deepest top to the highest, always including zero for the baseline.
function scaleOf(layers: Layer[]): ValueScale {
  let min = 0;
  let max = 0;
  for (const layer of layers) {
    layer.tops.forEach((top, i) => {
      min = Math.min(min, top, layer.floors[i]!);
      max = Math.max(max, top, layer.floors[i]!);
    });
  }
  return { min, max };
}

function pointX(frame: Frame, i: number): number {
  return pointAt(frame.axis, frame.plot, i);
}

// One series' fill: the pieces of every stretch between two points. A stretch
// reaches into the next one only while another follows it.
function layerPieces(data: ChartData, layer: Layer, frame: Frame): Piece[] {
  const series = data.series[layer.index]!;
  const base = valueY(0, frame.scale, frame.plot);
  const { tops } = layer;
  const out: Piece[] = [];
  for (let i = 0; i + 1 < tops.length; i += 1) {
    const from = tops[i]!;
    const to = tops[i + 1]!;
    out.push(
      ...slopePieces(
        {
          left: pointX(frame, i),
          right: pointX(frame, i + 1),
          from: valueY(from, frame.scale, frame.plot),
          to: valueY(to, frame.scale, frame.plot),
          base,
        },
        {
          id: `${String(layer.index)}.${String(i)}`,
          color: series.colors[i]!,
          seam: AREA_SEAM,
          first: i === 0,
          last: i + 2 >= tops.length,
        },
      ),
    );
  }
  return out;
}

// Every k-th point carries a label once a slot is narrower than the widest
// label to be drawn: the first point always does.
function labelStride(data: ChartData, layers: Layer[], slot: number): number {
  const widths = layers.flatMap((layer) => {
    const series = data.series[layer.index]!;
    if (!showsLabels(series)) return [];
    return series.labels
      .filter((text, i) => text !== "" && layer.has[i] === true)
      .map((text) => textWidth(text, LABEL_SIZE));
  });
  return Math.max(1, Math.ceil(Math.max(0, ...widths) / slot));
}

// The points of a layer that carry a label: all of them (thinned) while the
// series has data labels, else just the last one that holds a value.
function labelledPoints(
  layer: Layer,
  series: ChartSeries,
  stride: number,
): number[] {
  const held = layer.has.flatMap((has, i) => (has ? [i] : []));
  if (showsLabels(series)) return held.filter((i) => i % stride === 0);
  const last = held.at(-1);
  return last === undefined ? [] : [last];
}

// A plain area's label: above the line where there is room (below it for a
// value under zero), else inside the fill on the other side of the line.
// Returns its top edge.
function besideLine(value: number, frame: Frame): number {
  const { plot, scale } = frame;
  const y = valueY(value, scale, plot);
  const outside = value >= 0 ? y - LABEL_HEIGHT : y;
  const room =
    outside >= plot.top && outside + LABEL_HEIGHT <= plot.top + plot.height;
  if (room) return outside;
  const inside = value >= 0 ? y : y - LABEL_HEIGHT;
  return clamp(inside, plot.top, plot.top + plot.height - LABEL_HEIGHT);
}

// A stacked area's label: centred in the band of its series, and none where
// the band is too thin to hold it. Returns its top edge.
function inBand(layer: Layer, i: number, frame: Frame): number | null {
  const { plot, scale } = frame;
  const upper = valueY(layer.tops[i]!, scale, plot);
  const lower = valueY(layer.floors[i]!, scale, plot);
  if (Math.abs(lower - upper) < MIN_SEGMENT) return null;
  const middle = (upper + lower) / 2 - LABEL_HEIGHT / 2;
  return clamp(middle, plot.top, plot.top + plot.height - LABEL_HEIGHT);
}

// A label in the ink; areaLabels recolours it once everything has its place.
function labelOf(
  data: ChartData,
  layer: Layer,
  i: number,
  frame: Frame,
): Primitive[] {
  const text = data.series[layer.index]!.labels[i]!;
  if (text === "") return [];
  const { plot } = frame;
  const top =
    data.kind === "stackedArea"
      ? inBand(layer, i, frame)
      : besideLine(layer.tops[i]!, frame);
  if (top === null) return [];
  const width = Math.min(textWidth(text, LABEL_SIZE), plot.width);
  const left = clamp(
    pointX(frame, i) - width / 2,
    plot.left,
    plot.left + plot.width - width,
  );
  const name = `label ${String(layer.index)}.${String(i)}`;
  return [
    label(boxAt(left, top, width, LABEL_HEIGHT), text, data.ink, "c", name),
  ];
}

// The value labels over the fills. The end labels of plain areas without data
// labels are over the same point and would sit on one another where the areas
// end close together: they are stacked apart (chart-shapes-ends.ts). A stack's
// labels are each in the band of their own series and cannot meet. Only then,
// with every label where it stays, is its ink chosen against the fill that is
// really under it (chart-shapes-area-hit.ts).
function areaLabels(
  data: ChartData,
  layers: Layer[],
  frame: Frame,
  fills: Piece[],
): Primitive[] {
  const stride = labelStride(data, layers, frame.axis.step);
  const ends: End[] = [];
  const out = layers.flatMap((layer) => {
    const series = data.series[layer.index]!;
    const points = labelledPoints(layer, series, stride);
    if (!showsLabels(series) && points[0] !== undefined) {
      const column = points[0];
      ends.push({
        name: `label ${String(layer.index)}.${String(column)}`,
        column,
      });
    }
    return points.flatMap((i) => labelOf(data, layer, i, frame));
  });
  const placed =
    data.kind === "stackedArea" ? out : spreadEnds(out, ends, frame.plot);
  return inkOver(placed, fills, data.ink);
}

// An area or a stacked area as primitives: the title, the fills (a stack
// highest series first), the zero line, the value labels over them, the
// category names and the legend.
export function layoutArea(data: ChartData, box: Box): Primitive[] {
  const { title, plot, categoryBand, legendBand } = bands(data, box);
  const stacked = data.kind === "stackedArea";
  const layers = stacked ? stackedLayers(data) : plainLayers(data);
  const scale = scaleOf(layers);
  const frame: Frame = { plot, scale, axis: pointAxis(data, plot) };
  const painted = stacked ? [...layers].reverse() : layers;
  const out: Primitive[] = [];
  if (data.title !== null) {
    out.push(titleText(title, data.title, data.titleColor));
  }
  const fills = painted.flatMap((layer) => layerPieces(data, layer, frame));
  out.push(...fills);
  out.push(baselineLine(data.kind, data, plot, scale));
  out.push(...areaLabels(data, layers, frame, fills));
  if (categoryBand) {
    out.push(...axisLabels(data, plot, categoryBand));
  }
  if (legendBand) out.push(...legend(data, legendBand));
  return out;
}
