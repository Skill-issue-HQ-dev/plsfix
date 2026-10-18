// Which Excel chart type is which kind, and what the series of one chart add up
// to: one kind for them all, or a combo when their types differ or one sits on
// the secondary axis; the flags a series carries beyond its values; and the
// stack a banker's total line folds into. Pure: no Office.js.
// Invariant: a type with no shapes is no kind, so the picture says why; a type
// the host did not name is no type, never the word "null" in a sentence.

import type { ChartKind } from "../link/chart-model";
import type { ChartSeries, SeriesType } from "../link/chart-series";
import { comboIssue } from "../link/combo-rules";

// Excel reports every series of a tornado as fully overlapped.
export const FULL_OVERLAP = 100;

type Kind = Exclude<ChartKind, "combo">;

export interface Classified {
  kind: Kind;
  percent: boolean;
}

// First match wins, so the bar families are named before the clustered
// catch-all that would otherwise swallow "BarClustered".
const KIND_RULES: readonly (readonly [RegExp, Kind])[] = [
  [/^Waterfall$/, "waterfall"],
  [/^(3D)?Pie(Exploded)?$/, "pie"],
  [/^Doughnut(Exploded)?$/, "doughnut"],
  [/^Line(Markers)?$/, "line"],
  [/^(3D)?AreaStacked(100)?$/, "stackedArea"],
  [/^(3D)?Area$/, "area"],
  [/^XYScatter/, "scatter"],
  [/^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Stacked/, "stackedBar"],
  [/^(3D)?(Bar|CylinderBar|ConeBar|PyramidBar)Clustered$/, "bar"],
  [/^(3D)?(Column|Cylinder|Cone|Pyramid)(Col)?.*Stacked/, "stackedColumn"],
  [/Clustered$/, "column"],
];

// The kind an Excel chart (or series) type is drawn as, and whether it is a
// 100 % stack; null for every type that stays the picture it is today.
export function classify(chartType: string): Classified | null {
  const kind = KIND_RULES.find(([pattern]) => pattern.test(chartType))?.[1];
  if (kind === undefined) return null;
  return { kind, percent: /Stacked100$/.test(chartType) };
}

// Clustered columns laid out side by side would draw one bar behind another
// when the series fully overlap: that chart is one the picture says better.
export function overlapsColumns(
  chartType: string,
  overlap: number | null,
): boolean {
  return classify(chartType)?.kind === "column" && overlap === FULL_OVERLAP;
}

// The kind a slide can draw as shapes, or null for every chart type that stays
// the picture - and for columns that overlap (overlapsColumns).
export function chartKind(
  chartType: string,
  overlap: number | null,
): ChartKind | null {
  return overlapsColumns(chartType, overlap)
    ? null
    : (classify(chartType)?.kind ?? null);
}

// The type a host named, or "" when it named none. Excel for the web answers
// null for the type of a chart whose series differ (a combo; read on
// 05.10.2026): String(null) made that the type "null", and a sentence said so.
export function typeNamed(value: unknown): string {
  return typeof value === "string" ? value.trim() : "";
}

// A type the slide has no shapes for, named the way Excel names it; a type
// nobody named is not named at all, the sentence says what the user sees.
export function unsupportedType(chartType: string | null | undefined): string {
  const named = typeNamed(chartType);
  return named === ""
    ? "this chart type is not drawn as shapes"
    : `${named} charts are not drawn as shapes`;
}

// What one series said about itself. A host that did not say - a type that
// refuses the reads, an older one - leaves the defaults: no type, the primary
// axis, nothing known about labels and markers.
export interface SeriesFacts {
  chartType: string;
  axisGroup: string;
  hasDataLabels: boolean | null;
  markerStyle: string;
  // Whether a line series is drawn as a curve; read for a line chart only,
  // so absent or null is "not known" and flags nothing.
  smooth?: boolean | null;
}

export const UNKNOWN_FACTS: SeriesFacts = {
  chartType: "",
  axisGroup: "",
  hasDataLabels: null,
  markerStyle: "",
};

// What the series of a chart add up to: one kind for them all, or - when their
// types differ or they sit on both axes - a combo that names each series' own
// type and which of them are 100 % stacks (`percents`, a combo's only); or the
// reason there is no shape version.
export type Decision =
  | {
      kind: ChartKind;
      percent: boolean;
      types: SeriesType[] | null;
      percents?: boolean[];
    }
  | { issue: string };

// What a decision is when it is not the reason there is no shape version.
export type Decided = Exclude<Decision, { issue: string }>;

// The type that stands for the whole chart when it is asked about as one: its
// own, else - when the host named none - the first series' that did.
export function standingType(
  headType: string | null | undefined,
  facts: readonly SeriesFacts[],
): string {
  const named = typeNamed(headType);
  if (named !== "") return named;
  return facts.find((one) => one.chartType !== "")?.chartType ?? "";
}

// The chart's own type decides what a series that names none is drawn as. A
// host that names none for the chart (the web, for a combo) leaves it to the
// series: each is its own type, a series that names none is the first named
// one's, and a chart whose series name nothing either stays a picture.
export function decideKind(
  headType: string | null | undefined,
  facts: readonly SeriesFacts[],
): Decision {
  const named = typeNamed(headType);
  const head = named === "" ? null : classify(named);
  if (named !== "" && head === null) return { issue: unsupportedType(named) };
  if (head?.kind === "waterfall") {
    return { kind: "waterfall", percent: false, types: null };
  }
  const own = facts.map((one) =>
    one.chartType === "" ? null : classify(one.chartType),
  );
  const odd = facts.findIndex(
    (one, i) =>
      one.chartType !== "" && (own[i] === null || own[i]?.kind === "waterfall"),
  );
  if (odd >= 0) {
    return { issue: `${facts[odd]!.chartType} series are not drawn as shapes` };
  }
  const stands = head ?? own.find((one) => one !== null) ?? null;
  if (stands === null) return { issue: unsupportedType(null) };
  const typed: Classified[] = own.map((one) => one ?? stands);
  const first = typed[0] ?? stands;
  const mixed = typed.some(
    (one) => one.kind !== first.kind || one.percent !== first.percent,
  );
  // Series on two axes are two scales; all of them on the secondary one are one.
  const onSecondary = facts.map((one) => one.axisGroup === "Secondary");
  const twoAxes = onSecondary.includes(true) && onSecondary.includes(false);
  if (mixed || twoAxes) {
    const types = typed.map((one) => one.kind as SeriesType);
    const percents = typed.map((one) => one.percent);
    return { kind: "combo", percent: false, types, percents };
  }
  return { kind: first.kind, percent: first.percent, types: null };
}

// Whether the combo a decision names is one the slide can draw (combo-rules.ts):
// the reason it is not, or null - also for a decision that is no combo. A 100 %
// stack with a line is declined here: C2 kept it a combo so it would never be
// drawn as an absolute stack, and the combo layout has no 100 % stack to draw
// it with, so it is the picture, with that sentence.
export function comboVerdict(
  decision: Decided,
  facts: readonly SeriesFacts[],
): string | null {
  if (decision.kind !== "combo" || decision.types === null) return null;
  return comboIssue(
    decision.types.map((type, i) => ({
      type,
      secondary: facts[i]?.axisGroup === "Secondary",
      percent: decision.percents?.[i] === true,
    })),
  );
}

// The flags a series carries beyond its values: the secondary axis, no data
// labels, (for a line or a scatter) no markers and (for a line) a smooth curve.
// Each is written only when it differs from what an absent field means.
type Flags = Pick<ChartSeries, "axis" | "dataLabels" | "markers" | "smooth">;

export function seriesFlags(fact: SeriesFacts, type: SeriesType): Flags {
  const flags: Flags = {};
  if (fact.axisGroup === "Secondary") flags.axis = "secondary";
  if (fact.hasDataLabels === false) flags.dataLabels = false;
  const marked = type === "line" || type === "scatter";
  if (marked && fact.markerStyle === "None") flags.markers = false;
  if (type === "line" && fact.smooth === true) flags.smooth = true;
  return flags;
}

// One series as the total-folding rule sees it.
export interface TotalCandidate {
  name: string;
  type: SeriesType;
  values: (number | null)[];
  labels: string[];
  dataLabels: boolean;
  secondary: boolean;
  // A series of a 100 % stack: its sums are not what the chart plots.
  percent: boolean;
}

const SUM_TOLERANCE = 1e-9;

function sameAmount(a: number, b: number): boolean {
  return (
    Math.abs(a - b) <= SUM_TOLERANCE * Math.max(1, Math.abs(a), Math.abs(b))
  );
}

// A banker's total on top of a stack is a line series that sits on the stack's
// own sums with its labels on. Such a chart is the stack, and the line is its
// total: which series stay, and the label of each category's total. Null for
// any chart that is not exactly that, which stays a combo - a 100 % stack among
// them, since the folded chart would be an absolute stack the sheet never drew.
export function foldTotals(
  series: readonly TotalCandidate[],
): { keep: number[]; totals: string[] } | null {
  const lines = series.flatMap((one, i) => (one.type === "line" ? [i] : []));
  const at = lines[0];
  if (lines.length !== 1 || at === undefined) return null;
  const line = series[at]!;
  const keep = series.flatMap((_one, i) => (i === at ? [] : [i]));
  const stack = keep.map((i) => series[i]!);
  const plain = stack.every(
    (one) => one.type === "stackedColumn" && !one.secondary && !one.percent,
  );
  if (!plain || stack.length === 0 || !line.dataLabels || line.secondary) {
    return null;
  }
  if (line.values.every((value) => value === null)) return null;
  const sums = line.values.map((_value, i) =>
    stack.reduce((sum, one) => sum + (one.values[i] ?? 0), 0),
  );
  const onTheStack = line.values.every(
    (value, i) => value === null || sameAmount(value, sums[i]!),
  );
  if (!onTheStack) return null;
  const totals = line.labels.map((text, i) =>
    line.values[i] === null ? "" : text,
  );
  return { keep, totals };
}
