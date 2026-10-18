// One series of a chart payload: what it may carry (a gap as null, a type and
// an axis for a combo, label and marker flags, an x for a scatter), the
// defaults an absent field stands for, and the one reorder every per-point
// array of a chart goes through. Pure: no Office.js.
// Invariant: a payload from a pane that knew none of these fields reads
// exactly as that pane drew it.

import type { ChartData, ChartKind } from "./chart-model";

// What one series is drawn as: a chart's own kind for every series of it, or
// the kind a combo's series names for itself.
export type SeriesType =
  | "column"
  | "stackedColumn"
  | "bar"
  | "stackedBar"
  | "line"
  | "area"
  | "stackedArea"
  | "scatter"
  | "pie"
  | "doughnut";

export type SeriesAxis = "primary" | "secondary";

// One series: a value, its displayed text and its brand colour per category.
// A null value is a blank cell, a gap: a column, a bar, a line or a slice has
// nothing there, never a zero. An area is the one kind that plots it as a zero,
// as Excel does (chart-shapes-area.ts).
export interface ChartSeries {
  name: string;
  values: (number | null)[];
  labels: string[];
  colors: string[];
  // Absent = the chart's own kind; a combo names one per series.
  type?: SeriesType;
  // Absent = primary; a combo's margin line is typically secondary.
  axis?: SeriesAxis;
  // Excel's hasDataLabels and "marker style is not none"; absent = drawn,
  // which is how every pane before these fields drew every chart.
  dataLabels?: boolean;
  markers?: boolean;
  // Excel's `smooth` on a line series, written only when true. The slide has
  // no curve through points, so it draws the line straight and says so
  // (smoothNote in chart-model.ts); an older pane ignores the field.
  smooth?: boolean;
  // A scatter's X for each point, beside the Y in `values`.
  x?: (number | null)[];
}

const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const SERIES_TYPES: readonly string[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "line",
  "area",
  "stackedArea",
  "scatter",
  "pie",
  "doughnut",
];
const AXES: readonly string[] = ["primary", "secondary"];

// The series type a chart of this kind is made of. A waterfall draws its
// bars like a column; a combo has no type of its own, so a series that names
// none falls back to the commonest one.
const BASE_TYPE: Record<ChartKind, SeriesType> = {
  column: "column",
  stackedColumn: "stackedColumn",
  bar: "bar",
  stackedBar: "stackedBar",
  waterfall: "column",
  pie: "pie",
  line: "line",
  doughnut: "doughnut",
  area: "area",
  stackedArea: "stackedArea",
  scatter: "scatter",
  combo: "column",
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isList(value: unknown, points: number): value is unknown[] {
  return Array.isArray(value) && value.length === points;
}

// A number a slide can place, or null for a blank cell: nothing else travels.
function isGapValues(value: unknown, points: number): boolean {
  return (
    isList(value, points) &&
    value.every(
      (one) =>
        one === null || (typeof one === "number" && Number.isFinite(one)),
    )
  );
}

function isOptional(value: unknown, ok: (entry: unknown) => boolean): boolean {
  return value === undefined || ok(value);
}

export function isSeries(value: unknown, points: number): value is ChartSeries {
  if (!isRecord(value) || typeof value.name !== "string") return false;
  const { labels, colors } = value;
  return (
    isGapValues(value.values, points) &&
    isList(labels, points) &&
    labels.every((one) => typeof one === "string") &&
    isList(colors, points) &&
    colors.every((one) => typeof one === "string" && HEX_COLOR.test(one)) &&
    isOptional(value.x, (one) => isGapValues(one, points)) &&
    isOptional(value.type, (one) => SERIES_TYPES.includes(String(one))) &&
    isOptional(value.axis, (one) => AXES.includes(String(one))) &&
    isOptional(value.dataLabels, (one) => typeof one === "boolean") &&
    isOptional(value.markers, (one) => typeof one === "boolean") &&
    isOptional(value.smooth, (one) => typeof one === "boolean")
  );
}

// The type a series of a chart of this kind is drawn as.
export function baseType(kind: ChartKind): SeriesType {
  return BASE_TYPE[kind];
}

export function seriesType(data: ChartData, series: ChartSeries): SeriesType {
  return series.type ?? BASE_TYPE[data.kind];
}

export function isSecondary(series: ChartSeries): boolean {
  return series.axis === "secondary";
}

export function showsLabels(series: ChartSeries): boolean {
  return series.dataLabels !== false;
}

export function showsMarkers(series: ChartSeries): boolean {
  return series.markers !== false;
}

// The maths that cannot hold a gap (a waterfall's running total) reads it as
// the zero the cell would add.
export function filled(values: readonly (number | null)[]): number[] {
  return values.map((value) => value ?? 0);
}

// What a pie divides into slices: the positive values only, a gap and a
// negative one add nothing.
export function positiveTotal(values: readonly (number | null)[]): number {
  return values.reduce<number>(
    (sum, value) => sum + Math.max(0, value ?? 0),
    0,
  );
}

// The chart with every per-point array taken in `order` (new[k] = old[order[k]])
// and nothing else changed: the categories, each series' values, labels, colours
// and x, and the totals. This is the only list of them, so a reorder can never
// leave one behind.
export function reindexPoints(
  data: ChartData,
  order: readonly number[],
): ChartData {
  const at = <T>(list: readonly T[]): T[] => order.map((i) => list[i]!);
  const { totals, ...rest } = data;
  return {
    ...rest,
    categories: at(data.categories),
    series: data.series.map((one) => {
      const { x, ...kept } = one;
      return {
        ...kept,
        values: at(one.values),
        labels: at(one.labels),
        colors: at(one.colors),
        ...(x ? { x: at(x) } : {}),
      };
    }),
    ...(totals ? { totals: at(totals) } : {}),
  };
}
