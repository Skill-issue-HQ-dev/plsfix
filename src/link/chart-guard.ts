// The chart field of a picture payload: optional cargo, never a reason to
// throw the picture away. Owns that one field's validation on the way out of
// decodePayload, so model.ts's picture validator never sees it. Invariant: a
// picture keeps a chart the slide can draw, or none and a chartIssue saying so.

import { CHART_VERSION, isChartData, isChartKind } from "./chart-model";

// The reason a chart that travelled with the picture is not drawn: the panes
// print it after "as a picture", the way every other chartIssue is printed.
export const CHART_UNREADABLE = "chart data unreadable";

// A chart of a version this pane has never read, or whose only fault is a kind
// it has never heard of, was written by a newer pane: it is not broken, it is
// ahead, and the fix is an update.
export const CHART_NEEDS_UPDATE =
  "this chart needs a newer pls,fix; update pls,fix";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

// A version above the newest this pane reads, or an unknown kind and nothing
// else wrong (the same chart as a column would pass).
function isAhead(chart: unknown): boolean {
  if (!isRecord(chart)) return false;
  if (typeof chart.v === "number" && chart.v > CHART_VERSION) return true;
  return (
    typeof chart.kind === "string" &&
    !isChartKind(chart.kind) &&
    isChartData({ ...chart, kind: "column" })
  );
}

// Runs on the parsed JSON before the payload validator, because a chart the
// validator refuses would otherwise take the whole payload - picture and all -
// down with it. Anything but a picture carrying a broken chart is handed back
// untouched, the same object, so a good payload costs nothing.
export function guardChart(value: unknown): unknown {
  if (!isRecord(value) || value.kind !== "picture") return value;
  if (value.chart === undefined || isChartData(value.chart)) return value;
  const guarded: Record<string, unknown> = { ...value };
  delete guarded.chart;
  // The payload's own reason wins: Excel knows better than we do why the
  // chart it sent cannot be drawn.
  if (!("chartIssue" in guarded)) {
    guarded.chartIssue = isAhead(value.chart)
      ? CHART_NEEDS_UPDATE
      : CHART_UNREADABLE;
  }
  return guarded;
}
