// The chart data a picture link may carry: the kinds a slide can draw, the
// caps past which the picture is the honest answer and the words that say so,
// the validator every decoder runs, and the angle rule PowerPoint's pie
// adjustments follow. Every field a newer pane added is optional, so a payload
// from an older one still validates and reads as it always did.
// Pure: no Office.js, nothing from src/excel or src/ppt.

import { isSeries, type ChartSeries } from "./chart-series";

export type { ChartSeries };

// Every kind a payload may name: the one list, so a registry keyed by it
// cannot forget one and a test can walk them all.
export const CHART_KINDS = [
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
  "scatter",
  "combo",
] as const;

export type ChartKind = (typeof CHART_KINDS)[number];

export interface ChartData {
  // 1, or 2 for a 100 % stack: a pane from before them reads 1 only, refuses
  // the payload and keeps the picture, never drawing the stack as absolute.
  v: 1 | 2;
  kind: ChartKind;
  title: string | null;
  categories: string[];
  series: ChartSeries[];
  font: string;
  ink: string;
  titleColor: string;
  // A tornado: both series share one row, drawn longest first.
  overlap?: true;
  // A 100 % stack: each category's segments are drawn as shares of that
  // category (the values stay absolute, the way the sheet holds them).
  percent?: true;
  // The category axis' reversePlotOrder as the sheet had it. On a bar or a
  // stacked bar an ABSENT flag means top-down, the order every pane drew before
  // this field existed, so a payload already on a slide never flips; true is
  // top-down too, and only an explicit false is Excel's default (first category
  // at the bottom). On every other kind absent = false = first category first.
  categoriesReversed?: boolean;
  // The category axis is on its tick marks (Excel's "On tick marks", an area's
  // default): the first point is on the plot's left edge and the last on its
  // right. Written only when Excel says so, for a line or an area; absent is a
  // point at the centre of its slot, as every pane drew them before this field.
  onTicks?: true;
  // A stack's total over each category, as the sheet shows it.
  totals?: string[];
  // A doughnut's hole as Excel's doughnutHoleSize gives it: a whole percent,
  // 10 to 90, of the doughnut's own size (Microsoft's VBA and Office.js docs
  // say "a percentage of the chart size", OOXML's holeSize 10 to 90 the same):
  // the hole's diameter over the outer diameter, which is also the inner radius
  // over the outer one. Absent = what Excel draws a new doughnut with.
  hole?: number;
}

// What Excel gives a doughnut whose hole nobody set (Excel 2013 and later).
export const DOUGHNUT_HOLE_DEFAULT = 75;

// The hole as a fraction of the outer diameter: 0.55 for a 55 % hole.
export function holeFraction(data: ChartData): number {
  return (data.hole ?? DOUGHNUT_HOLE_DEFAULT) / PERCENT;
}

// The kinds drawn as slices of a whole, one slice per category: a pie and a
// doughnut (whose every series is a ring of slices).
export function isSliced(kind: ChartKind): boolean {
  return kind === "pie" || kind === "doughnut";
}

// Whether a bar or stacked bar is drawn the way Excel draws one whose axis is
// not reversed: first category (and, in a cluster, first series) at the bottom.
// Only an explicit false says so; absent or true is the top-down order panes
// drew before the flag existed. The one place that rule is written.
export function barsBottomUp(data: ChartData): boolean {
  return data.categoriesReversed === false;
}

// The highest chart version this pane reads; a chart above it is ahead.
export const CHART_VERSION = 2;
export const CHART_MAX_POINTS = 120;
export const CHART_MAX_SERIES = 12;
export const CHART_MAX_SLICES = 24;
export const CHART_MIN_POINTS = 2;
export const CHART_TITLE_MAX = 80;
export const DOUGHNUT_HOLE_MIN = 10;
export const DOUGHNUT_HOLE_MAX = 90;

// The kinds that can be a 100 % stack, and the ones that carry a total label.
const PERCENT_KINDS: readonly string[] = [
  "stackedColumn",
  "stackedBar",
  "stackedArea",
];
const TOTAL_KINDS: readonly string[] = ["stackedColumn", "stackedBar"];
// The kinds that draw points along a category axis, so can have it on its ticks.
export const ON_TICKS_KINDS: readonly ChartKind[] = [
  "line",
  "area",
  "stackedArea",
];
const FULL_TURN = 360;
const HALF_TURN = 180;
const PERCENT = 100;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStrings(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((v) => typeof v === "string");
}

function isReadVersion(value: unknown): value is 1 | 2 {
  return value === 1 || value === 2;
}

export function isChartKind(value: unknown): value is ChartKind {
  return CHART_KINDS.some((kind) => kind === value);
}

function withinTitle(title: unknown): boolean {
  return (
    title === null ||
    (typeof title === "string" && title.length <= CHART_TITLE_MAX)
  );
}

// What a kind needs of its series beyond the shared shape: a pie one series, a
// doughnut a ring of up to 24 slices for each of its series, a scatter an x on
// every series, a combo a type on each.
function fitsKind(
  kind: ChartKind,
  series: ChartSeries[],
  points: number,
): boolean {
  if (kind === "pie") {
    return series.length === 1 && points <= CHART_MAX_SLICES;
  }
  if (kind === "doughnut") return points <= CHART_MAX_SLICES;
  if (kind === "scatter") return series.every((one) => one.x !== undefined);
  if (kind === "combo") {
    return series.length >= 2 && series.every((one) => one.type !== undefined);
  }
  return true;
}

function holeFits(kind: ChartKind, hole: unknown): boolean {
  if (hole === undefined) return true;
  return (
    kind === "doughnut" &&
    typeof hole === "number" &&
    Number.isInteger(hole) &&
    hole >= DOUGHNUT_HOLE_MIN &&
    hole <= DOUGHNUT_HOLE_MAX
  );
}

// The optional chart-level fields: each is absent, or what the writer would
// have written and on a kind that can carry it.
function optionalsFit(
  kind: ChartKind,
  value: Record<string, unknown>,
  points: number,
): boolean {
  const { overlap, percent, categoriesReversed, totals, onTicks } = value;
  return (
    (overlap === undefined || overlap === true) &&
    (onTicks === undefined ||
      (onTicks === true && ON_TICKS_KINDS.includes(kind))) &&
    (percent === undefined ||
      (percent === true && PERCENT_KINDS.includes(kind))) &&
    (categoriesReversed === undefined ||
      typeof categoriesReversed === "boolean") &&
    (totals === undefined ||
      (TOTAL_KINDS.includes(kind) &&
        isStrings(totals) &&
        totals.length === points)) &&
    holeFits(kind, value.hole)
  );
}

// Shape, consistency and the caps in one pass: a payload that fails here
// simply has no chart data, and the picture beside it is what gets drawn.
export function isChartData(value: unknown): value is ChartData {
  if (!isRecord(value) || !isReadVersion(value.v) || !isChartKind(value.kind)) {
    return false;
  }
  if (!withinTitle(value.title) || !isStrings(value.categories)) return false;
  const points = value.categories.length;
  if (points < CHART_MIN_POINTS || points > CHART_MAX_POINTS) return false;
  const { series } = value;
  if (!Array.isArray(series) || series.length < 1) return false;
  if (series.length > CHART_MAX_SERIES) return false;
  if (!series.every((one) => isSeries(one, points))) return false;
  if (typeof value.font !== "string" || typeof value.ink !== "string") {
    return false;
  }
  if (typeof value.titleColor !== "string") return false;
  return (
    fitsKind(value.kind, series as ChartSeries[], points) &&
    optionalsFit(value.kind, value, points)
  );
}

// PowerPoint stores a pie's start and end as degrees clockwise from 3 o'clock
// and reads any angle back on (-180, 180]: 200 becomes -160, 360 becomes 0.
export function normalizeAngle(degrees: number): number {
  const wrapped =
    ((((degrees + HALF_TURN) % FULL_TURN) + FULL_TURN) % FULL_TURN) - HALF_TURN;
  return wrapped === -HALF_TURN ? HALF_TURN : wrapped;
}

// The reason a picture link has no chart data, in the words the panes show
// after "as a picture": the first cap the chart is outside of, counted the way
// the modeller sees the chart.
export function seriesCountIssue(count: number): string | null {
  if (count < 1) return "no series";
  if (count > CHART_MAX_SERIES) {
    return `${String(count)} series; shapes draw up to ${String(CHART_MAX_SERIES)}`;
  }
  return null;
}

export function chartCapIssue(
  kind: ChartKind,
  points: number,
  seriesCount: number,
): string | null {
  const series = seriesCountIssue(seriesCount);
  if (series !== null) return series;
  if (points < CHART_MIN_POINTS) {
    const noun = points === 1 ? "point" : "points";
    return `${String(points)} ${noun}; shapes need at least ${String(CHART_MIN_POINTS)}`;
  }
  if (points > CHART_MAX_POINTS) {
    return `${String(points)} points; shapes draw up to ${String(CHART_MAX_POINTS)}`;
  }
  if (kind === "pie" && seriesCount !== 1) {
    return `a pie with ${String(seriesCount)} series; pie shapes draw one`;
  }
  if (isSliced(kind) && points > CHART_MAX_SLICES) {
    return `${String(points)} slices; ${kind} shapes draw up to ${String(CHART_MAX_SLICES)}`;
  }
  return null;
}

// The sentence both panes show for a chart that stays a picture: the insert
// note in PowerPoint and the export line in Excel repeat it word for word.
const PICTURE_LEAD = "as a picture: ";
export function pictureNote(issue: string): string {
  return `${PICTURE_LEAD}${issue}`;
}

// A note without its "as a picture" lead: the bare reason, for a place where
// nothing is drawn at all (a Put back that is dead), so it names no picture.
export function pictureIssue(note: string): string {
  return note.startsWith(PICTURE_LEAD) ? note.slice(PICTURE_LEAD.length) : note;
}

// The only reasons left for a picture, in the words the row shows (D.1 of the
// 03.10.2026 plan): this host has no shape API at all (PowerPointApi below
// 1.8), the kind needs the newer one (pie, doughnut, area and stacked area,
// 1.10), or no shape version exists yet (kindDecline in chart-kinds.ts for a
// kind, comboIssue in combo-rules.ts for a combo made of what has none). A
// budget, a size, a cap, a blank and a stalled host are not reasons: they are
// drawn, or they say so.
export const NO_SHAPE_API = pictureNote(
  "this PowerPoint has no shape API (Microsoft 365 2504 / Mac 16.96 or newer)",
);

export function needsNewerPowerPoint(kind: string): string {
  return pictureNote(
    `${kind} shapes need Microsoft 365 2601 / Mac 16.105 or newer`,
  );
}

// What a line chart drawn as shapes could not keep of its source: Excel's
// smoothed curve through the points. The slide has no such shape, so the line is
// straight segments between the points, and the row says so after "Inserted"
// and under "Update all" (and Excel's own row after the export).
export const SMOOTH_STRAIGHT = "smooth lines are drawn straight";

export function smoothNote(data: ChartData): string | undefined {
  return data.series.some((one) => one.smooth === true)
    ? SMOOTH_STRAIGHT
    : undefined;
}

// What a stacked area (also 100 %) could not keep of a value under zero: the
// slide paints each series from the baseline up to its running total, which is
// the band between two totals only while no series is negative. Said on the
// same three rows as the smooth note.
export const STACK_NEGATIVE =
  "values under zero in a stack are drawn approximately";

export function stackNote(data: ChartData): string | undefined {
  const under = data.series.some((one) =>
    one.values.some((value) => value !== null && value < 0),
  );
  return data.kind === "stackedArea" && under ? STACK_NEGATIVE : undefined;
}

// The one note a drawn chart's row carries: what it could not keep of its
// source (a kind has at most one of the two).
export function drawnNote(data: ChartData): string | undefined {
  return smoothNote(data) ?? stackNote(data);
}

const DRAWN_NOTES = [SMOOTH_STRAIGHT, STACK_NEGATIVE];

// A note as it reads after its label: a drawn chart's note is a fragment, so it
// goes in brackets (what it was joined to follows the bracket); any other note
// follows as it stands.
function clauseOf(note: string): string {
  const drawn = DRAWN_NOTES.find((one) => note.startsWith(one));
  return drawn === undefined ? note : `(${drawn})${note.slice(drawn.length)}`;
}

// A repaint's note after its link's label, as "Update all" lists it. A drawn
// chart's note is a fragment, so it goes in brackets after the label, the way
// Excel's toast shows it; a sentence of its own ("as a picture: ...", the
// deck's copy was not saved) follows the label as it stands, and so does what
// the fragment was joined to.
export function labelledNote(label: string, note: string): string {
  return `${label} ${clauseOf(note)}`;
}

// Whether a note reads on from the label it follows ("as a picture: ...", a
// drawn chart's fragment) rather than being a sentence of its own.
function readsOn(note: string): boolean {
  return note.startsWith(PICTURE_LEAD) || clauseOf(note) !== note;
}

// What follows a link's label in a toast ("Inserted <label>", "Put <label>
// back"), full stop included, so the label and its notes read as one sentence:
// the note that reads on from the label goes before the full stop ("Inserted
// Combo chart as a picture: <reason>."), any other note is a sentence of its own
// after it ("Inserted Sales. Placed over other objects: ...").
export function labelTail(notes: readonly string[]): string {
  const given = notes.filter((note) => note !== "");
  const at = given.findIndex(readsOn);
  const clause = at < 0 ? "" : ` ${clauseOf(given[at]!)}`;
  const sentences = given.filter((_note, i) => i !== at);
  const stop = clause.endsWith(".") ? "" : ".";
  const after = sentences.length === 0 ? "" : ` ${sentences.join(". ")}`;
  return `${clause}${stop}${after}`;
}

// The picture the user asked for, either at the long-draw question or with an
// explicit picture export: nothing declined it, so nothing says "as a picture".
export const PICTURE_CHOSEN = "Picture (your choice)";

// What a stalled host and a cancelled question say. A swallowed insert takes
// its own shapes down again and leaves the export in the Inbox; a swallowed
// refresh never takes the old group down (it goes in the sync that groups).
export const INSERT_STOPPED =
  "PowerPoint stopped answering; nothing was inserted, the export is still in the Inbox";
export const REFRESH_STOPPED =
  "PowerPoint stopped answering; the old chart stays";
export const INSERT_CANCELLED =
  "Cancelled: nothing was inserted, the export is still in the Inbox";
