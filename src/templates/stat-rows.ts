// The statistic rows under a comps-style table: the six the product's own Comps
// stats writes (src/stats.ts), each behind a COUNT guard. Owns the formulas and
// where the median and quartiles sit. Invariant: a column with no number reads
// empty, never an error, while a genuine error in a figure still shows.

import { STATS_LABELS } from "../stats";
import { blank, calc, label, type TemplateCell } from "../template-cells";
import { span } from "./refs";

type StatLabel = (typeof STATS_LABELS)[number];

// One formula per label, keyed by it, so a label Comps stats gains without a
// formula here fails the typecheck instead of shipping a gap.
const FORMULAS: Record<StatLabel, (range: string) => string> = {
  Min: (range) => `MIN(${range})`,
  "25th percentile": (range) => `PERCENTILE.INC(${range},0.25)`,
  Median: (range) => `MEDIAN(${range})`,
  Mean: (range) => `AVERAGE(${range})`,
  "75th percentile": (range) => `PERCENTILE.INC(${range},0.75)`,
  Max: (range) => `MAX(${range})`,
};

/** How many rows the statistics take. */
export const STAT_COUNT = STATS_LABELS.length;

/** Offsets from the first statistic row of the rows a block prices off. */
export const STAT_ROW = {
  p25: STATS_LABELS.indexOf("25th percentile"),
  median: STATS_LABELS.indexOf("Median"),
  p75: STATS_LABELS.indexOf("75th percentile"),
} as const;

export interface StatSpec {
  /** Block row of the first and last data row the statistics read. */
  first: number;
  last: number;
  /** Block columns that carry a statistic; the label sits in column 0. */
  columns: number[];
  /** The block's width, every row of which this returns whole. */
  width: number;
}

// Empty when no peer has a number: a loss-making or blank column should not
// show #NUM! or #DIV/0! that the model check then has to explain.
function guarded(statistic: StatLabel, range: string): string {
  return `=IF(COUNT(${range})=0,"",${FORMULAS[statistic](range)})`;
}

/** The six statistic rows, `width` cells each, labels down column 0. */
export function statRows(spec: StatSpec): TemplateCell[][] {
  return STATS_LABELS.map((name) =>
    Array.from({ length: spec.width }, (_unused, column) => {
      if (column === 0) return label(name);
      if (!spec.columns.includes(column)) return blank();
      const range = span(spec.first, column, spec.last, column);
      return calc(guarded(name, range), "multiple");
    }),
  );
}
