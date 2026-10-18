// The registry of chart kinds: which layout draws each kind, which way its
// categories run (so the slide shows them in the order Excel does), and which
// kinds have no layout yet and decline with an honest sentence. The one place a
// kind is added: a layout for it, or `pending` until its slice lands.
// Pure: no Office.js. Invariant: every ChartKind has a rule.

import { layoutClassic, type Primitive } from "./chart-shapes";
import { layoutArea } from "./chart-shapes-area";
import { layoutCombo } from "./chart-shapes-combo";
import { layoutDoughnut } from "./chart-shapes-doughnut";
import { layoutStack } from "./chart-shapes-stack";
import { MIN_BRIDGE_POINTS } from "./chartmath";
import type { Box } from "./layout";
import {
  barsBottomUp,
  type ChartData,
  type ChartKind,
} from "./link/chart-model";
import { reindexPoints } from "./link/chart-series";
import { comboIssue, comboParts } from "./link/combo-rules";

// A layout turns a chart into primitives inside a box, or null when this kind
// has none (a kind with no layout yet, or a combo made of what has no shapes).
type Layout = (data: ChartData, box: Box) => Primitive[] | null;

interface Rule {
  layout: Layout;
  // Which way the categories run when the axis is not reversed: across (the
  // first category at the left, a column's, line's or area's axis) or up (the
  // first category at the bottom, a bar's - drawn so only where the payload
  // says its axis is not reversed, see barsBottomUp); null where there is no
  // category axis to reverse.
  axis: "across" | "up" | null;
  // The name a kind with no layout yet is declined under.
  pending?: string;
  // The PowerPointApi set the host needs to draw this kind as shapes: 1.8 for
  // groups, text boxes and the formats a chart wears, 1.10 where the shapes
  // need adjustments (a pie's angles, an area's triangle tips, a doughnut's
  // arcs) and rotation. Every drawn kind names one; a pending kind has none.
  api?: "1.8" | "1.10";
}

// Wrapped, not named, so the registry never reads a layout before the module
// that owns it has finished loading (chart-shapes.ts calls back into here).
const classic: Layout = (data, box) => layoutClassic(data, box);
const stacked: Layout = (data, box) => layoutStack(data, box);
const areas: Layout = (data, box) => layoutArea(data, box);
const rings: Layout = (data, box) => layoutDoughnut(data, box);
const combos: Layout = (data, box) => layoutCombo(data, box);
const notYet: Layout = () => null;

// A bridge reads an opening total, steps and a closing total, so a waterfall of
// fewer points than that has nothing to draw (bridgeSeries would throw): its
// reason, or null for one that can be drawn.
function bridgeDecline(data: ChartData): string | null {
  const points = data.categories.length;
  if (points >= MIN_BRIDGE_POINTS) return null;
  const noun = points === 1 ? "point" : "points";
  return `${String(points)} ${noun}; waterfall shapes need at least ${String(MIN_BRIDGE_POINTS)}`;
}
const bridges: Layout = (data, box) =>
  bridgeDecline(data) === null ? layoutClassic(data, box) : null;

const RULES: Record<ChartKind, Rule> = {
  column: { layout: classic, axis: "across", api: "1.8" },
  stackedColumn: { layout: stacked, axis: "across", api: "1.8" },
  bar: { layout: classic, axis: "up", api: "1.8" },
  stackedBar: { layout: stacked, axis: "up", api: "1.8" },
  waterfall: { layout: bridges, axis: null, api: "1.8" },
  pie: { layout: classic, axis: null, api: "1.10" },
  line: { layout: classic, axis: "across", api: "1.8" },
  doughnut: { layout: rings, axis: null, api: "1.10" },
  area: { layout: areas, axis: "across", api: "1.10" },
  stackedArea: { layout: areas, axis: "across", api: "1.10" },
  scatter: { layout: notYet, axis: null, pending: "scatter" },
  combo: { layout: combos, axis: "across", api: "1.8" },
};

// The chart with its categories in the order they are drawn, first to last
// from the left or the top: a bar chart whose payload says its axis is not
// reversed runs from the bottom, as Excel draws it (a payload with no flag is
// top-down, as panes always drew it), every other category axis runs from the
// first unless it is reversed; a chart in the opposite order is turned round.
export function displayOrder(data: ChartData): ChartData {
  const { axis } = RULES[data.kind];
  const reversed = data.categoriesReversed === true;
  const turn =
    axis === "up" ? barsBottomUp(data) : axis === "across" && reversed;
  if (!turn) return data;
  const last = data.categories.length - 1;
  return reindexPoints(
    data,
    data.categories.map((_name, i) => last - i),
  );
}

// The primitives of a chart, or null for a kind that has no layout yet.
export function layoutKind(data: ChartData, box: Box): Primitive[] | null {
  return RULES[data.kind].layout(displayOrder(data), box);
}

// Why a kind is a picture for now: the sentence after "as a picture", or null
// for a kind that is drawn.
export function kindDecline(kind: ChartKind): string | null {
  const { pending } = RULES[kind];
  return pending === undefined
    ? null
    : `${pending} charts are not drawn as shapes yet`;
}

// Why this chart is a picture, from the chart and not only its kind: the kind's
// sentence, or - for a combo made of what has no shapes - the series that has
// none, or - for a waterfall too short to bridge - its point count, in the
// words both panes print after "as a picture". Null for a chart that is drawn.
export function chartDecline(data: ChartData): string | null {
  const kind = kindDecline(data.kind);
  if (kind !== null) return kind;
  if (data.kind === "waterfall") return bridgeDecline(data);
  return data.kind === "combo" ? comboIssue(comboParts(data)) : null;
}

// A kind as a sentence names it: "stacked area", not "stackedArea".
export function kindWords(kind: ChartKind): string {
  return kind === "stackedArea" ? "stacked area" : kind;
}

// The PowerPointApi set the host needs to draw this kind as shapes, or
// undefined for a kind with no layout yet (kindDecline says why).
export function kindApi(kind: ChartKind): "1.8" | "1.10" | undefined {
  return RULES[kind].api;
}
