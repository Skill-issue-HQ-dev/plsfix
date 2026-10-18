// The horizontal bar and tornado-row primitives: one row per category, the
// series side by side within it or (a tornado) overlapped. Split out of
// chart-shapes.ts to keep both it and chart-shapes-parts.ts under the line cap.
// Pure: no Office.js. Invariant: every primitive lies inside the box it was given.

import {
  BAR_FILL,
  LABEL_HEIGHT,
  NO_CURSOR,
  valueX,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import type { Box } from "./layout";
import {
  barsBottomUp,
  type ChartData,
  type ChartSeries,
} from "./link/chart-model";
import { showsLabels } from "./link/chart-series";
import { barRange, boxAt, rect, segmentLabel } from "./chart-shapes-parts";

// The slot order for one category: input order, or (a tornado) the widest
// swing first regardless of sign, so the shorter bar behind it stays visible.
// `bottomUp` is Excel's own order for a bar chart whose axis the payload says is
// not reversed: the first series of a cluster sits lowest, so the slots run last
// to first. Without it (a payload from before the flag) they run first to last.
export function rowOrder(
  series: ChartSeries[],
  i: number,
  overlap: boolean,
  bottomUp = false,
): number[] {
  const indices = series.map((_, j) => j);
  if (!overlap) return bottomUp ? indices.reverse() : indices;
  return indices.sort(
    (a, b) =>
      Math.abs(series[b]!.values[i] ?? 0) - Math.abs(series[a]!.values[i] ?? 0),
  );
}

// Bar and tornado rows: one row per category, the series stacked vertically
// within it (or overlapping, longer first, when `overlap` is set). A blank
// cell leaves its place in the row empty. The categories arrive in display
// order; the series inside a row follow Excel's own order for the axis. Value
// labels are all dropped once a bar's share of its row is shorter than a label
// is high, the way a column chart drops them when its slot is narrower than
// the widest label.
export function barBars(
  data: ChartData,
  chart: Box,
  plot: Box,
  scale: ValueScale,
): Primitive[] {
  const overlap = data.overlap === true;
  const bottomUp = barsBottomUp(data);
  const rowSlot = plot.height / data.categories.length;
  const perRow = overlap ? 1 : data.series.length;
  const barHeight = (BAR_FILL * rowSlot) / perRow;
  const showValues = rowSlot / perRow >= LABEL_HEIGHT;
  const out: Primitive[] = [];
  data.categories.forEach((_, i) => {
    rowOrder(data.series, i, overlap, bottomUp).forEach((j, slotIndex) => {
      const series = data.series[j]!;
      const value = series.values[i];
      if (value === null || value === undefined) return;
      const [lo, hi] = barRange(false, NO_CURSOR, value);
      // A tornado row is one band; otherwise the series take a slot each.
      const top = overlap
        ? plot.top + i * rowSlot + (rowSlot - barHeight) / 2
        : plot.top +
          i * rowSlot +
          (rowSlot - data.series.length * barHeight) / 2 +
          slotIndex * barHeight;
      const left = valueX(lo, scale, plot);
      const box = boxAt(left, top, valueX(hi, scale, plot) - left, barHeight);
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      if (showValues && showsLabels(series)) {
        out.push(...segmentLabel(data, series, j, i, box, false, false, chart));
      }
    });
  });
  return out;
}
