// The combo kind: the column series (clustered, or stacked) on the primary
// scale exactly as the column layout draws them and the line series as the line
// layout draws them, each series on ITS axis - a series on the secondary axis
// has a scale of its own made from its own values (zero in it, like every scale
// here), so a 0-40 % margin line over columns of EUR 0-13 392 k uses the plot's
// full height. Both scales keep a label's room under the top of the plot, so no
// label of the tallest column or the highest point touches the title. A combo
// made of what has no shapes (combo-rules.ts) has no layout: null. Pure: no
// Office.js. Invariant: every primitive lies inside the box it was given.

import {
  baselineLine,
  bands,
  type Primitive,
  type ValueScale,
} from "./chart-shapes";
import { axisLabels } from "./chart-shapes-axis";
import { columnBars } from "./chart-shapes-column";
import { lineSeries, type LineScale } from "./chart-shapes-line";
import { legend, titleText } from "./chart-shapes-parts";
import { withRoom } from "./chart-shapes-room";
import {
  columnStacks,
  scaleOf as stackScale,
  segmentGrid,
  type Grid,
} from "./chart-shapes-stack";
import type { Box } from "./layout";
import type { ChartData } from "./link/chart-model";
import { comboParts, comboRoles, type ComboRoles } from "./link/combo-rules";

// A label's height (18) and the radius of the largest marker (3): what the
// highest point of either scale stands under the top of the plot, so the label
// over it stays out of the title band.
export const COMBO_HEADROOM = 21;

// The two scales of a combo: the primary one (columns and the lines on that
// axis) and, when a line sits on the secondary axis, that axis' own.
interface Scales {
  primary: ValueScale;
  secondary: ValueScale | null;
}

// A scale from zero over a set of values, widened by the values it is given: a
// blank cell reaches nowhere.
function reach(
  values: readonly (number | null)[],
  into: ValueScale = { min: 0, max: 0 },
): ValueScale {
  let { min, max } = into;
  for (const value of values) {
    if (value === null) continue;
    min = Math.min(min, value);
    max = Math.max(max, value);
  }
  return { min, max };
}

// The chart with only the stacked columns' series holding values: the stack
// layout reads series by their place and a blank cell is no segment, so the
// other series (lines) add nothing to a stack and keep every name's index.
function stacksOnly(data: ChartData, roles: ComboRoles): ChartData {
  const own = new Set(roles.columns);
  return {
    ...data,
    series: data.series.map((one, j) =>
      own.has(j) ? one : { ...one, values: one.values.map(() => null) },
    ),
  };
}

// The primary scale: what the columns reach (a stack's sums when stacked) and
// the lines on the primary axis.
function primaryOf(data: ChartData, roles: ComboRoles, grid: Grid | null) {
  const base =
    grid === null
      ? reach(roles.columns.flatMap((j) => data.series[j]!.values))
      : stackScale(grid);
  return reach(
    roles.primaryLines.flatMap((j) => data.series[j]!.values),
    base,
  );
}

function scalesOf(
  data: ChartData,
  roles: ComboRoles,
  grid: Grid | null,
  plot: Box,
): Scales {
  const secondary =
    roles.secondaryLines.length === 0
      ? null
      : withRoom(
          reach(roles.secondaryLines.flatMap((j) => data.series[j]!.values)),
          plot.height,
          COMBO_HEADROOM,
        );
  return {
    primary: withRoom(
      primaryOf(data, roles, grid),
      plot.height,
      COMBO_HEADROOM,
    ),
    secondary,
  };
}

// Each line on the scale of its own axis; none for a series that is a column.
function lineScale(roles: ComboRoles, scales: Scales): LineScale {
  return (index) => {
    if (roles.primaryLines.includes(index)) return scales.primary;
    return roles.secondaryLines.includes(index) ? scales.secondary : null;
  };
}

// A combo as primitives: the title, the columns with their value labels, the
// zero line over them, the lines over that with their markers and labels, the
// category names and the legend of every series in the payload's order. Null
// for a combo that has no shapes.
export function layoutCombo(data: ChartData, box: Box): Primitive[] | null {
  const roles = comboRoles(comboParts(data));
  if (roles === null) return null;
  const { title, plot, categoryBand, legendBand } = bands(data, box);
  const stacks = roles.stacked ? stacksOnly(data, roles) : null;
  const grid = stacks === null ? null : segmentGrid(stacks);
  const scales = scalesOf(data, roles, grid, plot);
  const out: Primitive[] = [];
  if (data.title !== null) {
    out.push(titleText(title, data.title, data.titleColor));
  }
  if (stacks !== null && grid !== null) {
    out.push(...columnStacks(stacks, box, plot, plot, scales.primary, grid));
  } else {
    out.push(...columnBars(data, box, plot, scales.primary, roles.columns));
  }
  out.push(baselineLine("column", data, plot, scales.primary));
  out.push(...lineSeries(data, plot, lineScale(roles, scales)));
  if (categoryBand) out.push(...axisLabels(data, plot, categoryBand));
  if (legendBand) out.push(...legend(data, legendBand));
  return out;
}
