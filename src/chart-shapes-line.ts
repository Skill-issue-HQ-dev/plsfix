// The native line chart's primitives: one small ellipse marker per point, a
// connector between consecutive points, and the point's own value centred
// above its marker (or below it where the label above would start over the
// plot's top edge, in the title band). A blank cell has no point and breaks
// the line; a series without markers draws none and one without data labels
// labels its last point alone (those end labels are stacked apart where lines
// end close together, see chart-shapes-ends.ts); past MARKER_MAX_POINTS the
// markers are off and past one label per slot the labels are thinned to every
// n-th. A smooth line is drawn straight: the layout never reads the flag (the
// slide's row says so, see chart-model.ts).
// Pure: no Office.js, no DOM. Invariant: every primitive lies inside the box
// it was given.

import { pointAt, pointAxis, type PointAxis } from "./chart-shapes-axis";
import { spreadEnds, type End } from "./chart-shapes-ends";
import { boxAt, label, lineBetween } from "./chart-shapes-parts";
import {
  LABEL_HEIGHT,
  LABEL_SIZE,
  textWidth,
  valueY,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";
import { showsLabels, showsMarkers } from "./link/chart-series";

const MARKER_MIN = 3;
const MARKER_MAX = 6;
const MARKER_SPACING_DIVISOR = 5;
// A line with more points than this is a line: a marker on each would be a
// dotted row, and each is a shape the host has to be asked for.
export const MARKER_MAX_POINTS = 24;
const CONNECTOR_WEIGHT = 1.5;
// How far above the plot's own top a label may start and still count as inside
// it: a label that starts higher would spill into the title band above the plot
// and flips below its marker instead.
const TOP_EDGE_EPSILON = 0.01;

interface Point {
  left: number;
  top: number;
}

// What every point of the chart is drawn with: the marker size, whether the
// chart has few enough points for markers at all, and how many points a label
// skips.
interface Style {
  marker: number;
  markers: boolean;
  stride: number;
}

// Point i sits where the axis puts it (chart-shapes-axis.ts): the centre of
// slot i, half a slot in from the edge, or on the plot's edge when Excel has the
// axis on its ticks. A blank cell has no position at all.
function positionsFor(
  series: ChartSeries,
  axis: PointAxis,
  plot: Box,
  scale: ValueScale,
): (Point | null)[] {
  return series.values.map((value, pointIndex) =>
    value === null
      ? null
      : {
          left: pointAt(axis, plot, pointIndex),
          top: valueY(value, scale, plot),
        },
  );
}

// A marker never runs past the plot: its centre follows the point, but the
// box itself is pulled back to the nearest edge once it would spill out.
function markerBoxAt(point: Point, marker: number, plot: Box): Box {
  return boxAt(
    Math.max(
      plot.left,
      Math.min(point.left - marker / 2, plot.left + plot.width - marker),
    ),
    Math.max(
      plot.top,
      Math.min(point.top - marker / 2, plot.top + plot.height - marker),
    ),
    marker,
    marker,
  );
}

// The point's value, centred on its marker: above it normally, below it when
// above would start over the plot's top edge - the marker flush with that edge,
// or no more than a label's height under it - where the title is.
function pointLabel(
  content: string,
  ink: string,
  point: Point,
  markerBox: Box,
  plot: Box,
  name: string,
): Primitive {
  const width = textWidth(content, LABEL_SIZE);
  const left = Math.max(
    plot.left,
    Math.min(point.left - width / 2, plot.left + plot.width - width),
  );
  const above = markerBox.top - LABEL_HEIGHT;
  const spills = above < plot.top - TOP_EDGE_EPSILON;
  const top = spills ? markerBox.top + markerBox.height : above;
  return label(boxAt(left, top, width, LABEL_HEIGHT), content, ink, "c", name);
}

// What a series is drawn into: the chart (for its ink), the plot and the style.
interface Frame {
  data: ChartData;
  plot: Box;
  style: Style;
}

// The last cell of a series that holds a value, or -1 for a series of blanks:
// the one point whose value a series without data labels still shows.
function lastPoint(positions: (Point | null)[]): number {
  for (let at = positions.length - 1; at >= 0; at -= 1) {
    if (positions[at]) return at;
  }
  return -1;
}

// One point's primitives: the connector joining it to the point right before
// it (none when that one is a blank cell or this is the first of a run), its
// marker and its value label, each only where the series and the chart ask.
function pointPrimitives(
  series: ChartSeries,
  seriesIndex: number,
  pointIndex: number,
  positions: (Point | null)[],
  frame: Frame,
  last: number,
): Primitive[] {
  const point = positions[pointIndex];
  if (!point) return [];
  const { data, plot, style } = frame;
  const out: Primitive[] = [];
  const previous = positions[pointIndex - 1];
  if (previous) {
    out.push(
      lineBetween(
        previous.left,
        previous.top,
        point.left,
        point.top,
        series.colors[pointIndex]!,
        CONNECTOR_WEIGHT,
        `line ${seriesIndex}.${pointIndex - 1}`,
      ),
    );
  }
  const markerBox = markerBoxAt(point, style.marker, plot);
  if (style.markers && showsMarkers(series)) {
    out.push({
      kind: "ellipse",
      box: markerBox,
      color: series.colors[pointIndex]!,
      name: `marker ${seriesIndex}.${pointIndex}`,
    });
  }
  const labelled = showsLabels(series)
    ? pointIndex % style.stride === 0
    : pointIndex === last;
  if (labelled) {
    const name = `label ${seriesIndex}.${pointIndex}`;
    const text = series.labels[pointIndex]!;
    out.push(pointLabel(text, data.ink, point, markerBox, plot, name));
  }
  return out;
}

// Every k-th point carries a label once a slot is narrower than the widest
// label that will be drawn: the first point always does, and no two labels
// sit closer than the widest one is wide. Only the series that are lines of
// this chart count (`drawn`).
function labelStride(
  data: ChartData,
  slot: number,
  drawn: (index: number) => boolean,
): number {
  const widths = data.series
    .filter((series, index) => drawn(index) && showsLabels(series))
    .flatMap((series) =>
      series.labels
        .filter((_text, i) => series.values[i] !== null)
        .map((text) => textWidth(text, LABEL_SIZE)),
    );
  return Math.max(1, Math.ceil(Math.max(0, ...widths) / slot));
}

// A native line chart becomes one editable connector per segment, a small
// marker per point (up to MARKER_MAX_POINTS points) and a value label per
// point (thinned when they would touch). Markers make individual values
// selectable in PowerPoint and avoid a line disappearing into a
// matching-colour slide background. The scale is every series' own, or - for a
// combo, whose lines sit on two axes and whose other series are not lines - a
// function of the series' place that answers its scale, or null for a series
// that is not a line of this chart: nothing is drawn for it and its labels do
// not count in how far apart the others sit.
export type LineScale = ValueScale | ((index: number) => ValueScale | null);

export function lineSeries(
  data: ChartData,
  plot: Box,
  scale: LineScale,
): Primitive[] {
  const points = data.categories.length;
  const axis = pointAxis(data, plot);
  const scaleOf = typeof scale === "function" ? scale : () => scale;
  const style: Style = {
    marker: Math.min(
      MARKER_MAX,
      Math.max(MARKER_MIN, axis.step / MARKER_SPACING_DIVISOR),
    ),
    markers: points <= MARKER_MAX_POINTS,
    stride: labelStride(data, axis.step, (index) => scaleOf(index) !== null),
  };
  const frame: Frame = { data, plot, style };
  const ends: End[] = [];
  const out = data.series.flatMap((series, seriesIndex) => {
    const own = scaleOf(seriesIndex);
    if (own === null) return [];
    const positions = positionsFor(series, axis, plot, own);
    const last = lastPoint(positions);
    if (last >= 0 && !showsLabels(series)) {
      ends.push({
        name: `label ${String(seriesIndex)}.${String(last)}`,
        column: last,
      });
    }
    return positions.flatMap((_point, pointIndex) =>
      pointPrimitives(series, seriesIndex, pointIndex, positions, frame, last),
    );
  });
  // The end labels stay in the plot, never on the title above it.
  return spreadEnds(out, ends, plot);
}
