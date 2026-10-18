// Reading a chart the way a slide has to redraw it: which kind it is, its
// title, every series' values, the label text the model itself shows and the
// brand colour of each point. Owns the Excel-to-ChartData mapping, and decides
// every string, so the slide side never formats a number. The per-series reads
// (type, axis, flags, X and Y, the axis' order) live in chart-read-series.ts.
// Invariant: a chart it cannot read whole is no chart at all: the picture
// beside it is the deck's, and the reason travels with it in the words the
// panes show.

import { pieColors, seriesColor, waterfallColors } from "../chart-colors";
import {
  chartCapIssue,
  isChartData,
  ON_TICKS_KINDS,
  seriesCountIssue,
  type ChartData,
  type ChartKind,
} from "../link/chart-model";
import {
  baseType,
  filled,
  type ChartSeries,
  type SeriesType,
} from "../link/chart-series";
import {
  activeTheme,
  getActiveSettings,
  type BrandSettings,
} from "../settings";
import {
  chartKind,
  comboVerdict,
  decideKind,
  foldTotals,
  FULL_OVERLAP,
  overlapsColumns,
  seriesFlags,
  standingType,
  typeNamed,
  unsupportedType,
  UNKNOWN_FACTS,
  type Decided,
  type SeriesFacts,
} from "./chart-read-kind";
import {
  padToLongest,
  readSeriesSet,
  valuesOf,
  type Dimensions,
  type SeriesSet,
} from "./chart-read-series";
import { formatChartAmount, hostSupports } from "./internal";
import { parseAddress } from "./shared";

export { chartKind, valuesOf };

// getDimensionValues: below it a chart link is the picture alone.
const DIMENSION_API = "1.12";
export const CHARTS_NEED_EXCEL_1_12 =
  "shape charts need Excel 2008/16.40 or newer";
export const SERIES_READ_REFUSED = "Excel refused the series read";
const OVERLAPPING_COLUMNS = "overlapping columns are not drawn as shapes";
const NOT_ALL_NUMBERS = "the series are not all numbers";
const OUTSIDE_LIMITS = "the chart is outside the shape limits";
// One area of A1, dollars or not: what worksheet.getRange takes.
const A1_RANGE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(:\$?[A-Za-z]{1,3}\$?\d{1,7})?$/;

// A chart the host names no type for (Excel for the web, for a combo) is asked
// for the reads of its commonest form, columns with lines on them: the series
// then say what they are, and decideKind decides from those.
const UNNAMED_CHART_READS: ChartKind = "column";

// The chart type ("" when the host named none) and title, read in the batch
// that carries the picture. `kind` is the kind the reads are queued for.
interface Head {
  chartType: string;
  kind: ChartKind;
  title: string | null;
  names: string[];
}

// What a read answers with: the data a slide draws as shapes, or the reason it
// has none, in the words the panes show after "as a picture".
export type ChartRead =
  { chart: ChartData; issue?: undefined } | { chart: null; issue: string };

function noChart(issue: string): ChartRead {
  return { chart: null, issue };
}

// What a chart link ships beside its picture, or the reason it ships nothing.
// Always commits at least one sync: the caller queues the picture into the
// same batch, so a chart nothing can be read from costs exactly the round trip
// a chart link costs today.
export async function readChartData(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<ChartRead> {
  if (!hostSupports(DIMENSION_API)) {
    await context.sync();
    return noChart(CHARTS_NEED_EXCEL_1_12);
  }
  const head = await readHead(context, chart);
  if (typeof head === "string") return noChart(head);
  try {
    return await readBody(context, chart, head);
  } catch {
    // A chart office.js will not describe - a type whose series refuse the
    // dimension reads, a source range that has gone - is one the deck draws
    // as the picture it already has.
    return noChart(SERIES_READ_REFUSED);
  }
}

// Sync one, shared with the picture: the type, the title and the series names.
// Everything else waits for the next batch, so a chart type that refuses to
// answer for it can never cost the picture its own round trip. A string is the
// reason there is no head to read on.
async function readHead(
  context: Excel.RequestContext,
  chart: Excel.Chart,
): Promise<Head | string> {
  chart.load("chartType");
  chart.title.load("text");
  chart.series.load("items/name");
  await context.sync();

  const names = chart.series.items.map((one) => String(one.name ?? ""));
  const count = seriesCountIssue(names.length);
  if (count !== null) return count;
  const chartType = typeNamed(chart.chartType);
  const kind =
    chartType === "" ? UNNAMED_CHART_READS : chartKind(chartType, null);
  if (kind === null) return unsupportedType(chartType);
  const text = String(chart.title.text ?? "").trim();
  return { chartType, kind, title: text === "" ? null : text, names };
}

// Syncs two and three: every series' dimensions and what the series say about
// themselves, then the cells the values came from. What the series add up to
// decides the kind: one kind for them all, or a combo.
async function readBody(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  head: Head,
): Promise<ChartRead> {
  const set = await readSeriesSet(context, chart, head);
  const standing = standingType(head.chartType, set.facts);
  if (overlapsColumns(standing, set.overlaps[0] ?? null)) {
    return noChart(OVERLAPPING_COLUMNS);
  }
  if (set.dimensions.some((one) => one === null)) {
    return noChart(NOT_ALL_NUMBERS);
  }
  const decision = decideKind(head.chartType, set.facts);
  if ("issue" in decision) return noChart(decision.issue);
  // A combo made of what the slide has no shapes for is the picture, with the
  // reason: drawn as an approximation it would be silently wrong.
  const refused = comboVerdict(decision, set.facts);
  if (refused !== null) return noChart(refused);
  const dimensions = set.dimensions as Dimensions[];
  const read =
    decision.kind === "scatter" ? padToLongest(dimensions) : dimensions;
  const labels = await readLabels(context, read);
  const data = assemble(head.title, read, labels, set, decision);
  if (data !== null) return { chart: data };
  const points = read[0]?.values.length ?? 0;
  return noChart(
    chartCapIssue(decision.kind, points, read.length) ?? OUTSIDE_LIMITS,
  );
}

// Sync three: the cells the values were read from, so a label reads exactly as
// the model shows it. A source that is not one plain range - a literal series,
// a defined name, a sheet that has gone - falls back to the house style, and so
// does any cell the range has nothing to say about. A blank cell has no label.
async function readLabels(
  context: Excel.RequestContext,
  series: Dimensions[],
): Promise<string[][]> {
  const house = (): string[][] =>
    series.map((one) => one.values.map(houseLabel));
  try {
    const ranges = series.map((one) => sourceRange(context, one.source));
    if (ranges.every((range) => range === null)) return house();
    await context.sync();
    return series.map((one, index) => labelsOf(one.values, ranges[index]));
  } catch {
    return house();
  }
}

function houseLabel(value: number | null): string {
  return value === null ? "" : formatChartAmount(value);
}

function sourceRange(
  context: Excel.RequestContext,
  source: string,
): Excel.Range | null {
  const { sheet, address } = parseAddress(source);
  if (sheet === "" || !A1_RANGE.test(address)) return null;
  const range = context.workbook.worksheets.getItem(sheet).getRange(address);
  range.load("text");
  return range;
}

function labelsOf(
  values: (number | null)[],
  range: Excel.Range | null | undefined,
): string[] {
  const text = range ? range.text.flat() : [];
  return values.map((value, index) => {
    const cell = text[index];
    if (value === null) return "";
    return cell === undefined || cell === "" ? formatChartAmount(value) : cell;
  });
}

// One series as the payload carries it: values, labels, colours, and the
// fields a newer pane reads (its own type in a combo, its axis, its flags, an
// X for a scatter). Every field that equals its absent default is left out.
function toSeries(
  one: Dimensions,
  labels: string[],
  decided: Decided,
  index: number,
  facts: SeriesFacts,
  settings: BrandSettings,
): ChartSeries {
  const type: SeriesType = decided.types?.[index] ?? baseType(decided.kind);
  return {
    name: one.name,
    values: one.values,
    labels,
    colors: colorsOf(decided.kind, one.values, index, settings),
    ...(decided.types ? { type } : {}),
    ...seriesFlags(facts, type),
    ...(one.x ? { x: one.x } : {}),
  };
}

// Categories, colours and the brand of the pane, checked against the caps the
// slide draws within: anything the validator refuses ships as no chart at all.
function assemble(
  title: string | null,
  read: Dimensions[],
  labels: string[][],
  set: SeriesSet,
  decided: Decided,
): ChartData | null {
  const settings = getActiveSettings();
  const points = read[0]?.values.length ?? 0;
  const first = read[0]?.categories ?? [];
  const all = read.map((one, index) =>
    toSeries(
      one,
      labels[index] ?? [],
      decided,
      index,
      set.facts[index] ?? UNKNOWN_FACTS,
      settings,
    ),
  );
  const folded = foldedStack(all, decided);
  const kind = folded ? "stackedColumn" : decided.kind;
  const tornado =
    kind === "bar" &&
    set.overlaps.length > 0 &&
    set.overlaps.every((one) => one === FULL_OVERLAP);
  const data: ChartData = {
    v: decided.percent ? 2 : 1,
    kind,
    title,
    categories: first.length > 0 ? first : numbered(points),
    series: folded ? folded.series : all,
    font: settings.font,
    ink: activeTheme().formulaFont,
    titleColor: settings.primary,
    ...(tornado ? { overlap: true as const } : {}),
    ...(decided.percent ? { percent: true as const } : {}),
    ...categoryOrder(kind, set.reversed),
    ...tickPosition(kind, set.between),
    ...(folded ? { totals: folded.totals } : {}),
    ...(kind === "doughnut" && set.hole !== null ? { hole: set.hole } : {}),
  };
  return isChartData(data) ? data : null;
}

// The category axis' order as the payload carries it. A bar or stacked bar says
// what the sheet said, true or false, because a pane reads an absent flag as the
// top-down order of decks made before the flag; any other kind writes it only
// when the axis is reversed. Nothing at all where the axis could not be read.
function categoryOrder(
  kind: ChartKind,
  reversed: boolean | null,
): Pick<ChartData, "categoriesReversed"> {
  if (reversed === null) return {};
  if (kind === "bar" || kind === "stackedBar") {
    return { categoriesReversed: reversed };
  }
  return reversed ? { categoriesReversed: true } : {};
}

// The category axis on its tick marks, written only when Excel says so and for
// a kind that draws points along it (the validator refuses it on any other):
// absent is the centre of a slot, as every pane drew a line before the flag.
function tickPosition(
  kind: ChartKind,
  between: boolean | null,
): Pick<ChartData, "onTicks"> {
  return between === false && ON_TICKS_KINDS.includes(kind)
    ? { onTicks: true }
    : {};
}

// A stack with a banker's total line over it is the stack and its totals: the
// line goes, its labels become the totals, and what is left of the series is a
// plain stack (the per-series types a combo carries are dropped).
function foldedStack(
  all: ChartSeries[],
  decided: Decided,
): { series: ChartSeries[]; totals: string[] } | null {
  if (decided.kind !== "combo") return null;
  const folded = foldTotals(
    all.map((one, index) => ({
      name: one.name,
      type: one.type ?? "column",
      values: one.values,
      labels: one.labels,
      dataLabels: one.dataLabels !== false,
      secondary: one.axis === "secondary",
      percent: decided.percents?.[index] === true,
    })),
  );
  if (folded === null) return null;
  const series = folded.keep.map((at) => {
    const plain = { ...all[at]! };
    delete plain.type;
    return plain;
  });
  return { series, totals: folded.totals };
}

// The rules the sheet's own charts already follow: a waterfall by position and
// sign, a pie or doughnut by slice, everything else one palette colour per
// series (the palette's own six, then the same six shaded).
function colorsOf(
  kind: ChartKind,
  values: (number | null)[],
  index: number,
  settings: BrandSettings,
): string[] {
  if (kind === "waterfall") return waterfallColors(filled(values), settings);
  if (kind === "pie" || kind === "doughnut") {
    return pieColors(values.length, settings);
  }
  const color = seriesColor(index, settings);
  return values.map(() => color);
}

// A chart with no categories of its own numbers its points, the way Excel's own
// axis does.
function numbered(points: number): string[] {
  return Array.from({ length: points }, (_unused, index) => String(index + 1));
}
