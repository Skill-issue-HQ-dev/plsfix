// Where the points of a line or an area sit along the category axis, and the
// names under them. Between the ticks (a line's default in Excel) point i is at
// the centre of slot i; on them (an area's default, ChartData.onTicks) the first
// point is on the plot's left edge and the last on its right, n - 1 steps apart.
// Pure: no Office.js. Invariant: a name's box lies inside the plot's width.

import {
  categoryLabels,
  keptIndices,
  LABEL_SIZE,
  thinningFactor,
  widestWidth,
  type Text,
} from "./chart-shapes";
import { boxAt, label, truncate } from "./chart-shapes-parts";
import type { Box } from "./layout";
import { ON_TICKS_KINDS, type ChartData } from "./link/chart-model";

// A name on the ticks sits against the plot's edge at either end instead of
// centred on its tick, so it needs half as much room again before the axis
// leaves names out.
const END_ROOM = 1.5;

// Point i is at `(i + shift) * step` from the plot's left edge.
export interface PointAxis {
  shift: 0 | 0.5;
  step: number;
}

// Only a line or an area carries the flag, and a lone point has no ticks to be on.
function onTicks(data: ChartData): boolean {
  return (
    data.onTicks === true &&
    ON_TICKS_KINDS.includes(data.kind) &&
    data.categories.length > 1
  );
}

export function pointAxis(data: ChartData, plot: Box): PointAxis {
  const n = data.categories.length;
  return onTicks(data)
    ? { shift: 0, step: plot.width / (n - 1) }
    : { shift: 0.5, step: plot.width / n };
}

export function pointAt(axis: PointAxis, plot: Box, i: number): number {
  return plot.left + (i + axis.shift) * axis.step;
}

// The category names: the slots' own between the ticks; on them each centred on
// its tick, a box pulled in by the plot's edge keeping its name against that
// edge, by its tick.
export function axisLabels(data: ChartData, plot: Box, band: Box): Text[] {
  if (!onTicks(data)) return categoryLabels(data, plot, band, false);
  const axis = pointAxis(data, plot);
  const room = widestWidth(data.categories) * END_ROOM;
  const k = thinningFactor(axis.step, room);
  const width = Math.min(k * axis.step, plot.width);
  return keptIndices(data.categories.length, k).map((i) => {
    const centred = pointAt(axis, plot, i) - width / 2;
    const left = Math.max(
      plot.left,
      Math.min(centred, plot.left + plot.width - width),
    );
    const align = left > centred ? "l" : left < centred ? "r" : "c";
    const content = truncate(data.categories[i]!, LABEL_SIZE, width);
    const box = boxAt(left, band.top, width, band.height);
    return label(box, content, data.ink, align, `category ${i}`);
  });
}
