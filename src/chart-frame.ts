// The layout box a chart is redrawn in so that its drawing stands where the old
// one stood: a group's box is the union of its shapes, smaller than the box they
// were laid out in. Pure geometry over layoutChart, no Office.js.
// Invariant: the drawing at the frame has the reach's corner and width (+-0.5 pt).

import { layoutChart } from "./chart-shapes";
import type { Box } from "./layout";
import type { ChartData } from "./link/chart-model";

// What a drawing occupies: its left edge, top edge and width. An untitled chart
// leaves its title band empty, a column chart keeps room for its labels above
// the plot, a pie is as wide as its widest legend row: none fills its box.
export interface Reach {
  left: number;
  top: number;
  width: number;
}

// Half a point: a host reports a group's width in its own rounding, and the
// widths a layout can take are whole steps apart (a legend row holds one item
// more), so no step is ever mistaken for the target.
const REACH_TOLERANCE = 0.5;
// How wide, as a multiple of the reach, a box is tried at before giving up.
const MAX_GROWTH = 8;
const BISECTIONS = 40;
// Short of the first width that overshoots, so the chosen box sits inside its
// plateau and not on its edge.
const EDGE_MARGIN = 0.001;

// The union of what a layout draws, the way a host groups it: lines of no
// thickness count, like every other shape.
function reachOf(data: ChartData, box: Box): Reach | null {
  const boxes = layoutChart(data, box).map((primitive) => primitive.box);
  if (boxes.length === 0) return null;
  const left = Math.min(...boxes.map((one) => one.left));
  const top = Math.min(...boxes.map((one) => one.top));
  const right = Math.max(...boxes.map((one) => one.left + one.width));
  return { left, top, width: right - left };
}

// The smallest width in [low, high] for which `holds` is true, given that it
// is false at `low`, true at `high` and never turns false again after that.
function firstWhere(
  holds: (width: number) => boolean,
  low: number,
  high: number,
): number {
  let lo = low;
  let hi = high;
  for (let i = 0; i < BISECTIONS; i += 1) {
    const mid = (lo + hi) / 2;
    if (holds(mid)) hi = mid;
    else lo = mid;
  }
  return hi;
}

// The box width whose drawing is `target` wide. Most kinds fill their box
// edge to edge, so the target is the answer; a pie or a doughnut is as wide as
// its widest legend row, which grows a whole item at a time, so many box widths
// draw the same reach: the one nearest `hint` (the chart's own natural width)
// is taken, which is the width an untouched chart was inserted at.
function boxWidth(
  drawn: (width: number) => number,
  target: number,
  hint: number,
): number {
  const first = drawn(target);
  if (Math.abs(first - target) <= REACH_TOLERANCE) return target;
  if (first > target) return target;
  const reaches = (width: number): boolean =>
    drawn(width) >= target - REACH_TOLERANCE;
  const over = (width: number): boolean =>
    drawn(width) > target + REACH_TOLERANCE;
  const cap = target * MAX_GROWTH;
  let high = target * 2;
  while (!reaches(high) && high < cap) high *= 2;
  if (!reaches(high)) return target;
  const low = firstWhere(reaches, target, high);
  let top = Math.max(high, low * 2);
  while (!over(top) && top < cap) top *= 2;
  const edge = over(top) ? firstWhere(over, low, top) : top;
  const wanted = Number.isFinite(hint) ? hint : low;
  return Math.max(low, Math.min(Math.max(wanted, low), edge - EDGE_MARGIN));
}

// The box a rebuild lays the chart out in: `reach` is where the drawing must
// stand (what the host reports for the group), `aspect` the height a chart is
// given per point of width, `natural` the width the chart asks for before any
// slide shrinks it. A reach with no width, or a layout that draws nothing, is
// taken as the box itself.
export function frameFor(
  data: ChartData,
  reach: Reach,
  aspect: number,
  natural: number,
): Box {
  const sized = (width: number): Box => ({
    left: 0,
    top: 0,
    width,
    height: Math.round(width * aspect),
  });
  const plain: Box = {
    ...sized(reach.width),
    left: reach.left,
    top: reach.top,
  };
  if (!(reach.width > 0) || reachOf(data, sized(reach.width)) === null) {
    return plain;
  }
  const width = boxWidth(
    (one) => reachOf(data, sized(one))?.width ?? 0,
    reach.width,
    natural,
  );
  const drawn = reachOf(data, sized(width));
  if (drawn === null) return plain;
  return {
    ...sized(width),
    left: reach.left - drawn.left,
    top: reach.top - drawn.top,
  };
}
