// Reading the series of a chart from office.js: every series' dimensions
// (categories and values, or X and Y for a scatter) with a blank cell kept as a
// gap, what each series says about itself (its own type, axis group, data
// labels and marker style), the category axis' order and tick position, the
// overlap that marks a tornado and a doughnut's hole, all in the one batch the
// dimensions ride.
// Invariant: a host that refuses the new reads answers exactly what it answered
// before them, never a worse chart.

import {
  DOUGHNUT_HOLE_MAX,
  DOUGHNUT_HOLE_MIN,
  ON_TICKS_KINDS,
  type ChartKind,
} from "../link/chart-model";
import { UNKNOWN_FACTS, type SeriesFacts } from "./chart-read-kind";
import { formatChartAmount, hostSupports } from "./internal";

// ChartSeries.chartType, hasDataLabels, markerStyle and smooth are 1.7, axisGroup 1.8;
// the category axis' reversePlotOrder is 1.7, its isBetweenCategories 1.8;
// overlap is 1.8; the source string of a dimension is 1.15.
const FACTS_API = "1.8";
const AXIS_API = "1.7";
const OVERLAP_API = "1.8";
const SOURCE_API = "1.15";
// An error value Excel plots as an empty cell.
const NOT_AVAILABLE = "#N/A";
// The kinds whose first category sits on a category axis the user can reverse.
const CATEGORY_AXIS_KINDS: readonly ChartKind[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "line",
  "area",
  "stackedArea",
];

function isBlank(cell: unknown): boolean {
  if (cell === null || cell === undefined) return true;
  if (typeof cell === "number") return Number.isNaN(cell);
  if (typeof cell !== "string") return false;
  const text = cell.trim().toUpperCase();
  return text === "" || text === NOT_AVAILABLE;
}

// A dimension comes back as whatever the host felt like - numbers on the web,
// strings on other builds, an empty string or nothing for a blank cell - so
// every cell goes through Number(). A blank cell (and #N/A, which Excel plots
// as one) is a gap in the data, null, never a zero. Anything else that is not a
// number makes the whole series, and so the chart, a picture.
export function valuesOf(raw: readonly unknown[]): (number | null)[] | null {
  const values: (number | null)[] = [];
  for (const cell of raw) {
    if (isBlank(cell)) {
      values.push(null);
      continue;
    }
    const value = Number(cell);
    if (!Number.isFinite(value)) return null;
    values.push(value);
  }
  return values;
}

// One series' dimension reads, queued; the source string only on a host that
// names it, which is what decides whether labels cost a third sync.
interface Queued {
  // A scatter has X and Y and no categories: it is not asked for them.
  categories: OfficeExtension.ClientResult<string[]> | null;
  values: OfficeExtension.ClientResult<string[]>;
  x: OfficeExtension.ClientResult<string[]> | null;
  source: OfficeExtension.ClientResult<string> | null;
}

// What one series answered with, before any of it is a label or a colour.
export interface Dimensions {
  name: string;
  categories: string[];
  values: (number | null)[];
  x: (number | null)[] | null;
  source: string;
}

// What the body read adds to the head: every series' dimensions (null for one
// that is not all numbers), what each said about itself, the overlap that
// marks a tornado, the category axis' order and whether a line or an area has
// it between its tick marks (each null when it was not read: a host below the
// ExcelApi that has it, a kind with no such axis, or a read the host refused)
// and a doughnut's hole.
export interface SeriesSet {
  dimensions: (Dimensions | null)[];
  facts: SeriesFacts[];
  overlaps: number[];
  reversed: boolean | null;
  between: boolean | null;
  hole: number | null;
}

interface Wants {
  overlap: boolean;
  facts: boolean;
  axis: boolean;
  hole: boolean;
  // A line chart's series say whether they are curves.
  smooth: boolean;
  // A line or an area's category axis says whether it is between its ticks.
  between: boolean;
}

function wantsFor(kind: ChartKind): Wants {
  return {
    overlap: hostSupports(OVERLAP_API) && (kind === "column" || kind === "bar"),
    // The chartex types answer for none of it.
    facts: hostSupports(FACTS_API) && kind !== "waterfall",
    axis: hostSupports(AXIS_API) && CATEGORY_AXIS_KINDS.includes(kind),
    hole: kind === "doughnut",
    smooth: hostSupports(FACTS_API) && kind === "line",
    between: hostSupports(FACTS_API) && ON_TICKS_KINDS.includes(kind),
  };
}

function queueSeries(
  chart: Excel.Chart,
  index: number,
  scatter: boolean,
): Queued {
  const series = chart.series.getItemAt(index);
  const yName = scatter ? "YValues" : "Values";
  return {
    categories: scatter ? null : series.getDimensionValues("Categories"),
    values: series.getDimensionValues(yName),
    x: scatter ? series.getDimensionValues("XValues") : null,
    source: hostSupports(SOURCE_API)
      ? series.getDimensionDataSourceString(yName)
      : null,
  };
}

// Everything the body asks of the host, queued in one batch: the dimensions,
// then the loads. The extras (the series' own facts, the axis, the hole) ride
// the same batch as the dimensions so a working host costs no round trip more.
function queueBody(
  chart: Excel.Chart,
  names: readonly string[],
  kind: ChartKind,
  wants: Wants,
): Queued[] {
  const scatter = kind === "scatter";
  const queued = names.map((_name, index) =>
    queueSeries(chart, index, scatter),
  );
  const loads = [
    wants.overlap ? "items/overlap" : "",
    wants.facts
      ? "items/chartType,items/axisGroup,items/hasDataLabels,items/markerStyle"
      : "",
    wants.hole ? "items/doughnutHoleSize" : "",
    wants.smooth ? "items/smooth" : "",
  ].filter((path) => path !== "");
  if (loads.length > 0) chart.series.load(loads.join(","));
  const axis = chart.axes.categoryAxis;
  if (wants.axis) axis.load("reversePlotOrder");
  if (wants.between) axis.load("isBetweenCategories");
  return queued;
}

function text(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function factsOf(
  chart: Excel.Chart,
  count: number,
  wants: Wants,
): SeriesFacts[] {
  if (!wants.facts) return Array.from({ length: count }, () => UNKNOWN_FACTS);
  return chart.series.items.map((one) => ({
    chartType: text(one.chartType),
    axisGroup: text(one.axisGroup),
    hasDataLabels:
      typeof one.hasDataLabels === "boolean" ? one.hasDataLabels : null,
    markerStyle: text(one.markerStyle),
    // Read only where it was loaded: an unloaded scalar throws.
    smooth: wants.smooth && typeof one.smooth === "boolean" ? one.smooth : null,
  }));
}

// What the category axis answered for one of its flags, or null when it did not
// answer with a boolean.
function axisFlag(
  chart: Excel.Chart,
  name: "reversePlotOrder" | "isBetweenCategories",
): boolean | null {
  const flag: unknown = chart.axes.categoryAxis[name];
  return typeof flag === "boolean" ? flag : null;
}

// Excel's doughnutHoleSize of the first series, as the whole percent it is
// (10 to 90, "a percentage of the chart size": the hole's diameter over the
// doughnut's), carried unchanged; null when it was not read or is no such
// number, which the layout draws as Excel's default hole (chart-model.ts).
function holeOf(chart: Excel.Chart, wants: Wants): number | null {
  if (!wants.hole) return null;
  const hole = Number(chart.series.items[0]?.doughnutHoleSize);
  const fits =
    Number.isInteger(hole) &&
    hole >= DOUGHNUT_HOLE_MIN &&
    hole <= DOUGHNUT_HOLE_MAX;
  return fits ? hole : null;
}

function dimensionsOf(result: Queued, name: string): Dimensions | null {
  const values = valuesOf(result.values.value);
  const x = result.x ? valuesOf(result.x.value) : null;
  if (values === null || (result.x !== null && x === null)) return null;
  const categories = x
    ? x.map((one) => (one === null ? "" : formatChartAmount(one)))
    : (result.categories?.value ?? []).map(String);
  return {
    name,
    categories,
    values,
    x,
    source: result.source ? String(result.source.value ?? "") : "",
  };
}

// Series of one scatter need not be the same length (a two-point fit line beside
// the cloud): every series is padded with gaps to the longest, and they share
// that one's X as the categories, which is how Excel's blanks read.
export function padToLongest(series: Dimensions[]): Dimensions[] {
  const longest = series.reduce<Dimensions | null>(
    (best, one) =>
      best === null || one.values.length > best.values.length ? one : best,
    null,
  );
  if (longest === null) return series;
  const points = longest.values.length;
  const pad = <T>(list: T[], fill: T): T[] => [
    ...list,
    ...new Array<T>(Math.max(0, points - list.length)).fill(fill),
  ];
  return series.map((one) => ({
    ...one,
    categories: longest.categories,
    values: pad<number | null>(one.values, null),
    x: one.x ? pad<number | null>(one.x, null) : null,
  }));
}

// A refused batch asked again with less: without the two reads only a line or
// an area has (the smooth flag, the tick position: the other facts stay), and
// when that is refused too, without any extra.
async function askAgain(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  head: { kind: ChartKind; names: readonly string[] },
  wants: Wants,
): Promise<{ wants: Wants; queued: Queued[] }> {
  if (wants.smooth || wants.between) {
    const lean = { ...wants, smooth: false, between: false };
    const queued = queueBody(chart, head.names, head.kind, lean);
    try {
      await context.sync();
      return { wants: lean, queued };
    } catch {
      // The host refuses more than those two reads: ask for the dimensions.
    }
  }
  const bare = {
    ...wants,
    facts: false,
    axis: false,
    hole: false,
    smooth: false,
    between: false,
  };
  const queued = queueBody(chart, head.names, head.kind, bare);
  await context.sync();
  return { wants: bare, queued };
}

// Syncs two and three of a chart read, as one: every series' dimensions and,
// beside them, what the series say about themselves, the category axis' order
// and a doughnut's hole - asked for only where they mean something. A batch the
// host refuses while it carries those extras (a property it does not implement
// for this chart type, whatever code it says so with) is asked again: first
// without the smooth flag and the tick position, the extras of a line or an area
// a host may lack while it has the rest, then for the dimensions alone, which is
// exactly what it answered before they were asked; a chart that refuses that
// too is no chart.
export async function readSeriesSet(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  head: { kind: ChartKind; names: readonly string[] },
): Promise<SeriesSet> {
  let wants = wantsFor(head.kind);
  let queued = queueBody(chart, head.names, head.kind, wants);
  try {
    await context.sync();
  } catch (error) {
    const asked = wants.smooth || wants.facts || wants.axis || wants.between;
    if (!(asked || wants.hole)) throw error;
    ({ wants, queued } = await askAgain(context, chart, head, wants));
  }
  return {
    dimensions: queued.map((one, index) =>
      dimensionsOf(one, head.names[index] ?? ""),
    ),
    facts: factsOf(chart, head.names.length, wants),
    overlaps: wants.overlap
      ? chart.series.items.map((one) => Number(one.overlap))
      : [],
    reversed: wants.axis ? axisFlag(chart, "reversePlotOrder") : null,
    between: wants.between ? axisFlag(chart, "isBetweenCategories") : null,
    hole: holeOf(chart, wants),
  };
}
