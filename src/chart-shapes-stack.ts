// The stacked family: stacked and 100 % stacked, columns and bars. One
// rectangle per segment per category, positives up (or right) from the
// baseline and negatives below (or left of) it; a 100 % stack is each
// category's segments as shares of its own absolute total (the values stay as
// the sheet holds them); a blank cell leaves its slot empty; a segment label
// sits centred inside when it fits, a total label on top of (or at the end of)
// the stack when the chart carries one. Pure: no Office.js.
// Invariant: every primitive lies inside the box it was given.

import {
  LABEL_HEIGHT,
  baselineLine,
  bands,
  categoryLabels,
  BAR_FILL,
  textWidth,
  valueX,
  valueY,
  widestWidth,
  LABEL_SIZE,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import {
  boxAt,
  label,
  legend,
  rect,
  segmentLabel,
  stackSegment,
  titleText,
} from "./chart-shapes-parts";
import type { Box } from "./layout";
import type { ChartData } from "./link/chart-model";
import { showsLabels } from "./link/chart-series";

// One segment of one stack, as the stretch of the value axis it covers.
export interface Segment {
  lo: number;
  hi: number;
}

// Every segment of the chart, by category and then by series; null where the
// cell was blank.
export type Grid = (Segment | null)[][];

function clamp(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(value, high));
}

// Each category's segments in value space: the positives piled up from zero,
// the negatives down from it, in series order. A 100 % stack divides each
// value by its category's absolute total first, so a category of mixed signs
// still adds up to one whole and a category of zeros stays zero tall.
export function segmentGrid(data: ChartData): Grid {
  const percent = data.percent === true;
  return data.categories.map((_name, i) => {
    const values = data.series.map((one) => one.values[i] ?? null);
    const whole = percent
      ? values.reduce<number>((sum, value) => sum + Math.abs(value ?? 0), 0)
      : 1;
    const cursor = { pos: 0, neg: 0 };
    return values.map((value) => {
      if (value === null) return null;
      const share = percent ? (whole > 0 ? value / whole : 0) : value;
      const [lo, hi] = stackSegment(cursor, share);
      return { lo, hi };
    });
  });
}

// From the deepest negative stack to the tallest positive one, always
// including zero for the baseline.
export function scaleOf(grid: Grid): ValueScale {
  let min = 0;
  let max = 0;
  for (const segment of grid.flat()) {
    if (segment === null) continue;
    min = Math.min(min, segment.lo);
    max = Math.max(max, segment.hi);
  }
  return { min, max };
}

// Where a category's positive stack ends: what its total label sits against.
function topOf(row: (Segment | null)[]): number {
  return Math.max(0, ...row.map((segment) => segment?.hi ?? 0));
}

// The labels a stack's segments will carry, for the width rule.
function segmentTexts(data: ChartData): string[] {
  return data.series
    .filter(showsLabels)
    .flatMap((one) => one.labels.filter((_text, i) => one.values[i] !== null));
}

// The totals the chart draws: none for a 100 % stack, which has nothing to add
// up to, and none when no category has one.
function totalsShown(data: ChartData): string[] | null {
  if (data.percent === true || data.totals === undefined) return null;
  return data.totals.some((text) => text !== "") ? data.totals : null;
}

// The part of the plot the values are scaled over: the plot less the room a
// total label needs past the longest stack, so a label never leaves the box.
function valuePlotOf(plot: Box, horizontal: boolean, room: number): Box {
  if (horizontal) return { ...plot, width: Math.max(0, plot.width - room) };
  return {
    ...plot,
    top: plot.top + room,
    height: Math.max(0, plot.height - room),
  };
}

// Stacked columns: a slot per category, the series one on top of the next.
// Series outer so a category's segments (same left) sit next to each other in
// the output. A slot too narrow for the widest segment label gets none, the
// way a clustered column's does.
export function columnStacks(
  data: ChartData,
  chart: Box,
  plot: Box,
  valuePlot: Box,
  scale: ValueScale,
  grid: Grid,
): Primitive[] {
  const slot = plot.width / data.categories.length;
  const barWidth = BAR_FILL * slot;
  const showValues = slot >= widestWidth(segmentTexts(data));
  const out: Primitive[] = [];
  data.series.forEach((series, j) => {
    data.categories.forEach((_name, i) => {
      const segment = grid[i]![j];
      if (!segment) return;
      const top = valueY(segment.hi, scale, valuePlot);
      const box = boxAt(
        plot.left + i * slot + (slot - barWidth) / 2,
        top,
        barWidth,
        valueY(segment.lo, scale, valuePlot) - top,
      );
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      if (showValues && showsLabels(series)) {
        out.push(...segmentLabel(data, series, j, i, box, true, true, chart));
      }
    });
  });
  return out;
}

// Stacked bars: a row per category, the series one after the next within it,
// categories outer so a row's segments sit next to each other in the output.
function rowStacks(
  data: ChartData,
  chart: Box,
  plot: Box,
  valuePlot: Box,
  scale: ValueScale,
  grid: Grid,
): Primitive[] {
  const rowSlot = plot.height / data.categories.length;
  const barHeight = BAR_FILL * rowSlot;
  // Rows closer together than a label is high get no segment labels, the way
  // columns narrower than the widest label get none.
  const showValues = rowSlot >= LABEL_HEIGHT;
  const out: Primitive[] = [];
  data.categories.forEach((_name, i) => {
    data.series.forEach((series, j) => {
      const segment = grid[i]![j];
      if (!segment) return;
      const left = valueX(segment.lo, scale, valuePlot);
      const box = boxAt(
        left,
        plot.top + i * rowSlot + (rowSlot - barHeight) / 2,
        valueX(segment.hi, scale, valuePlot) - left,
        barHeight,
      );
      out.push(rect(box, series.colors[i]!, `bar ${j}.${i}`));
      if (showValues && showsLabels(series)) {
        out.push(...segmentLabel(data, series, j, i, box, true, false, chart));
      }
    });
  });
  return out;
}

// A total over each column, centred on it, thinned to every n-th once the
// columns are narrower than the widest total; the first always stays.
function columnTotals(
  data: ChartData,
  totals: string[],
  plot: Box,
  valuePlot: Box,
  scale: ValueScale,
  grid: Grid,
): Primitive[] {
  const slot = plot.width / totals.length;
  const stride = Math.max(1, Math.ceil(widestWidth(totals) / slot));
  const out: Primitive[] = [];
  totals.forEach((text, i) => {
    if (text === "" || i % stride !== 0 || grid[i]!.every((one) => !one))
      return;
    const width = Math.min(textWidth(text, LABEL_SIZE), plot.width);
    const centre = plot.left + (i + 0.5) * slot;
    const left = clamp(
      centre - width / 2,
      plot.left,
      plot.left + plot.width - width,
    );
    const top = valueY(topOf(grid[i]!), scale, valuePlot) - LABEL_HEIGHT;
    out.push(
      label(
        boxAt(left, top, width, LABEL_HEIGHT),
        text,
        data.ink,
        "c",
        `total ${i}`,
      ),
    );
  });
  return out;
}

// A total at the end of each row, thinned to every n-th row once the rows are
// closer together than a label is high.
function rowTotals(
  data: ChartData,
  totals: string[],
  plot: Box,
  valuePlot: Box,
  scale: ValueScale,
  grid: Grid,
): Primitive[] {
  const rowSlot = plot.height / totals.length;
  const stride = Math.max(1, Math.ceil(LABEL_HEIGHT / rowSlot));
  const out: Primitive[] = [];
  totals.forEach((text, i) => {
    if (text === "" || i % stride !== 0 || grid[i]!.every((one) => !one))
      return;
    const width = Math.min(textWidth(text, LABEL_SIZE), plot.width);
    const left = clamp(
      valueX(topOf(grid[i]!), scale, valuePlot),
      plot.left,
      plot.left + plot.width - width,
    );
    const top = clamp(
      plot.top + i * rowSlot + (rowSlot - LABEL_HEIGHT) / 2,
      plot.top,
      plot.top + plot.height - LABEL_HEIGHT,
    );
    out.push(
      label(
        boxAt(left, top, width, LABEL_HEIGHT),
        text,
        data.ink,
        "l",
        `total ${i}`,
      ),
    );
  });
  return out;
}

// A stacked or 100 % stacked chart as primitives: the title, the segments (and
// the totals), the zero line, the category names and the legend, in that order.
export function layoutStack(data: ChartData, box: Box): Primitive[] {
  const { title, plot, categoryBand, legendBand, labelColumn } = bands(
    data,
    box,
  );
  const horizontal = data.kind === "stackedBar";
  const grid = segmentGrid(data);
  const scale = scaleOf(grid);
  const totals = totalsShown(data);
  const room =
    totals === null ? 0 : horizontal ? widestWidth(totals) : LABEL_HEIGHT;
  const valuePlot = valuePlotOf(plot, horizontal, room);
  const out: Primitive[] = [];
  if (data.title !== null) {
    out.push(titleText(title, data.title, data.titleColor));
  }
  const stacks = horizontal ? rowStacks : columnStacks;
  out.push(...stacks(data, box, plot, valuePlot, scale, grid));
  if (totals !== null) {
    const writeTotals = horizontal ? rowTotals : columnTotals;
    out.push(...writeTotals(data, totals, plot, valuePlot, scale, grid));
  }
  out.push(baselineLine(data.kind, data, valuePlot, scale));
  if (horizontal && labelColumn) {
    out.push(...categoryLabels(data, plot, labelColumn, true));
  }
  if (categoryBand)
    out.push(...categoryLabels(data, plot, categoryBand, false));
  if (legendBand) out.push(...legend(data, legendBand));
  return out;
}
