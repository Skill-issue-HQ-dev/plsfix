// Clustered columns, split out of chart-shapes.ts at its line cap: one slot per
// category, the series side by side in it, drawn for the whole chart (the
// column kind) or for some of its series (a combo's columns, chart-shapes-
// combo.ts). Pure: no Office.js. Invariant: every primitive lies inside the
// box it was given, and a series' shapes are named by its place in the payload.

import {
  BAR_FILL,
  NO_CURSOR,
  valueY,
  widestWidth,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import { barRange, boxAt, rect, segmentLabel } from "./chart-shapes-parts";
import type { Box } from "./layout";
import type { ChartData, ChartSeries } from "./link/chart-model";
import { showsLabels } from "./link/chart-series";

// A slot too narrow for the widest value label gets no value labels at all: a
// column asking for what cannot fit is the "1 / 519" wrap this layout no longer
// produces.
function valuesFit(
  data: ChartData,
  plot: Box,
  group: readonly ChartSeries[],
): boolean {
  const slot = plot.width / data.categories.length;
  return slot >= widestWidth(group.flatMap((s) => s.labels));
}

// Whether a value label stands above any column of `members`: the labels fit
// their slot, a series shows them, and a value is not under zero (a negative
// column's label hangs below it). The value scale keeps room for them.
export function labelsAboveBars(
  data: ChartData,
  plot: Box,
  members: readonly number[] = data.series.map((_series, j) => j),
): boolean {
  const group = members.map((j) => data.series[j]!);
  return (
    valuesFit(data, plot, group) &&
    group.some(
      (s) =>
        showsLabels(s) &&
        s.values.some((value) => value !== null && value >= 0),
    )
  );
}

// Clustered columns: one slot per category, the series of `members` (every
// series by default) side by side in it; a blank cell leaves its place empty.
// Value labels follow valuesFit. The shapes of series j are named `bar j.i` and
// `label j.i` whichever place it takes in the slot.
export function columnBars(
  data: ChartData,
  chart: Box,
  plot: Box,
  scale: ValueScale,
  members: readonly number[] = data.series.map((_series, j) => j),
): Primitive[] {
  if (members.length === 0) return [];
  const slot = plot.width / data.categories.length;
  const group = members.map((j) => data.series[j]!);
  const showValues = valuesFit(data, plot, group);
  const barWidth = (BAR_FILL * slot) / group.length;
  const out: Primitive[] = [];
  members.forEach((j, place) => {
    const series = data.series[j]!;
    data.categories.forEach((_, i) => {
      const value = series.values[i];
      if (value === null || value === undefined) return;
      const left =
        plot.left +
        i * slot +
        (slot - group.length * barWidth) / 2 +
        place * barWidth;
      const [lo, hi] = barRange(false, NO_CURSOR, value);
      const top = valueY(hi, scale, plot);
      const box = boxAt(left, top, barWidth, valueY(lo, scale, plot) - top);
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      if (showValues && showsLabels(series)) {
        out.push(...segmentLabel(data, series, j, i, box, false, true, chart));
      }
    });
  });
  return out;
}
