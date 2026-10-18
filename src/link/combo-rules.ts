// Which combo charts a slide can draw, and what each series of one is for: the
// column series (clustered, or stacked) on the primary axis, the lines on
// whichever axis they say. Every other mix - a bar, an area or a scatter among
// the series, a 100 % stack, columns on the secondary axis, clustered and
// stacked columns together - has no shapes, and says so in one sentence both
// panes print after "as a picture". Pure: no Office.js. Invariant: a combo
// this file lets through is one chart-shapes-combo.ts draws exactly as Excel
// lays it out, never as an approximation.

import type { ChartData } from "./chart-model";
import { isSecondary, seriesType, type SeriesType } from "./chart-series";

// One series as the rules see it: what it is drawn as, which axis it says it is
// on, and whether it is a 100 % stack (only the Excel read knows: the payload
// of a combo never carries it, because such a combo is declined at the source).
export interface ComboPart {
  type: SeriesType;
  secondary: boolean;
  percent: boolean;
}

// What the layout needs: the places of the series in the payload, by role.
export interface ComboRoles {
  columns: number[];
  stacked: boolean;
  primaryLines: number[];
  secondaryLines: number[];
}

// A series type as a sentence names it.
const TYPE_WORDS: Readonly<Record<SeriesType, string>> = {
  column: "column",
  stackedColumn: "stacked column",
  bar: "bar",
  stackedBar: "stacked bar",
  line: "line",
  area: "area",
  stackedArea: "stacked area",
  scatter: "scatter",
  pie: "pie",
  doughnut: "doughnut",
};

function isColumn(type: SeriesType): boolean {
  return type === "column" || type === "stackedColumn";
}

// With nothing on the primary axis, the secondary one is the only axis the
// chart has: every series is on it, and so on the primary scale of the layout.
function axisOf(parts: readonly ComboPart[]): boolean[] {
  const anyPrimary = parts.some((one) => !one.secondary);
  return parts.map((one) => one.secondary && anyPrimary);
}

// The reason this combo has no shape version, or null when it has one. The
// first series that cannot be drawn names itself; then the mixes of columns.
export function comboIssue(parts: readonly ComboPart[]): string | null {
  const odd = parts.find((one) => !isColumn(one.type) && one.type !== "line");
  if (odd !== undefined) {
    return `${TYPE_WORDS[odd.type]} series in a combo chart are not drawn as shapes`;
  }
  if (parts.some((one) => one.percent)) {
    return "a 100 % stack in a combo chart is not drawn as shapes";
  }
  const types = new Set(
    parts.filter((one) => isColumn(one.type)).map((one) => one.type),
  );
  if (types.size > 1) {
    return "clustered and stacked columns in one combo chart are not drawn as shapes";
  }
  const secondary = axisOf(parts);
  if (parts.some((one, i) => isColumn(one.type) && secondary[i])) {
    return "columns on the secondary axis are not drawn as shapes";
  }
  return null;
}

// The series of a drawable combo by role, in payload order; null for a combo
// that declines.
export function comboRoles(parts: readonly ComboPart[]): ComboRoles | null {
  if (comboIssue(parts) !== null) return null;
  const secondary = axisOf(parts);
  const at = (keep: (part: ComboPart, onSecondary: boolean) => boolean) =>
    parts.flatMap((one, i) => (keep(one, secondary[i]!) ? [i] : []));
  return {
    columns: at((one) => isColumn(one.type)),
    stacked: parts.some((one) => one.type === "stackedColumn"),
    primaryLines: at((one, onSecondary) => one.type === "line" && !onSecondary),
    secondaryLines: at(
      (one, onSecondary) => one.type === "line" && onSecondary,
    ),
  };
}

// A payload's series as parts. A combo is never a 100 % stack here: the Excel
// read declines those, and the validator refuses the flag on a combo.
export function comboParts(data: ChartData): ComboPart[] {
  return data.series.map((one) => ({
    type: seriesType(data, one),
    secondary: isSecondary(one),
    percent: data.percent === true,
  }));
}
