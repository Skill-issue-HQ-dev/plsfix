// The scenario-switch block: a typed 1, 2 or 3 picks Base, Upside or Downside for
// four assumptions through CHOOSE into a Live column, and a three-year projection
// to unlevered FCF reads only that column. Owns the grid and its demo case; the
// invariant is that an unselected scenario cannot move a number.

import {
  blank,
  blankRow,
  build,
  calc,
  head,
  input,
  label,
  pad,
  result,
  type Template,
  type TemplateCell,
  title,
  total,
} from "../template-cells";
import { at } from "./refs";

const COLUMNS = 5;
// In the order the selector counts them: 1 = Base, 2 = Upside, 3 = Downside.
const SCENARIOS = ["Base", "Upside", "Downside"];
// Columns: the labels in 0, one per scenario, then the Live column.
const SCENARIO_COLUMNS = [1, 2, 3];
const LIVE = 4;
// The projection's columns: the base year in 1, then years one to three.
const PROJECTED_YEARS = [2, 3, 4];

// Where each line sits: the selector, the four assumptions, the projection.
const ROW = {
  selector: 1,
  growth: 5,
  margin: 6,
  capex: 7,
  tax: 8,
  revenue: 11,
  ebitda: 12,
  capexSpend: 13,
  taxPaid: 14,
  fcf: 15,
} as const;

const SELECTOR = at(ROW.selector, 1);

// The demo case: a base-year revenue of 100 EUR m and, per scenario, the growth,
// EBITDA margin, capex as a share of revenue and tax rate.
const DEMO = {
  baseRevenue: 100,
  growth: [0.06, 0.1, 0.02],
  margin: [0.2, 0.23, 0.16],
  capex: [0.08, 0.07, 0.1],
  tax: [0.2, 0.18, 0.22],
};

// The selector, the name of the scenario it reads as, and the key to its numbers.
function selectorRows(): TemplateCell[][] {
  const names = SCENARIOS.map((name) => `"${name}"`).join(",");
  const key = SCENARIOS.map((name, index) => `${index + 1} = ${name}`);
  return [
    pad(
      [
        label("Scenario (1, 2 or 3)"),
        input(1, "whole"),
        calc(`=CHOOSE(${SELECTOR},${names})`, "whole"),
      ],
      COLUMNS,
    ),
    pad([label(key.join(", "))], COLUMNS),
  ];
}

// One assumption: the three typed scenarios, then the Live column that picks the
// one the selector names. A selector outside 1 to 3 is #VALUE!, which is louder
// than quietly falling back on a scenario.
function assumptionRow(
  name: string,
  row: number,
  values: number[],
): TemplateCell[] {
  const picks = SCENARIO_COLUMNS.map((column) => at(row, column)).join(",");
  return [
    label(name),
    ...values.map((value) => input(value, "percent")),
    calc(`=CHOOSE(${SELECTOR},${picks})`, "percent"),
  ];
}

const header = (names: string[]): TemplateCell[] =>
  names.map((name) => head(name));

// The projection's revenue: a typed base year, then each year the one before
// grown by the live rate.
function revenueRow(): TemplateCell[] {
  return [
    label("Revenue (EUR m)"),
    input(DEMO.baseRevenue, "decimal"),
    ...PROJECTED_YEARS.map((column) =>
      calc(
        `=${at(ROW.revenue, column - 1)}*(1+${at(ROW.growth, LIVE)})`,
        "decimal",
      ),
    ),
  ];
}

// A projected line: nothing in the base year's column, a formula for each year.
function projected(
  name: string,
  formula: (column: number) => string,
  subtotal = false,
): TemplateCell[] {
  const first = subtotal ? result(name) : label(name);
  const gap = blank(subtotal ? "result" : "label");
  const cell = (column: number): TemplateCell =>
    subtotal
      ? total(formula(column), "decimal")
      : calc(formula(column), "decimal");
  return [first, gap, ...PROJECTED_YEARS.map(cell)];
}

// Tax is charged on a positive EBITDA less capex only, the capex standing in for
// the D&A that shields it, so a loss year pays none.
function projection(): TemplateCell[][] {
  const live = (row: number): string => at(row, LIVE);
  return [
    revenueRow(),
    projected(
      "EBITDA (EUR m)",
      (c) => `=${at(ROW.revenue, c)}*${live(ROW.margin)}`,
    ),
    projected(
      "Capex (EUR m)",
      (c) => `=${at(ROW.revenue, c)}*${live(ROW.capex)}`,
    ),
    projected("Tax on EBITDA less capex (EUR m)", (c) => {
      const cover = `${at(ROW.ebitda, c)}-${at(ROW.capexSpend, c)}`;
      return `=MAX(${cover},0)*${live(ROW.tax)}`;
    }),
    projected(
      "Unlevered FCF (EUR m, no working capital)",
      (c) =>
        `=${at(ROW.ebitda, c)}-${at(ROW.capexSpend, c)}-${at(ROW.taxPaid, c)}`,
      true,
    ),
  ];
}

export function scenarioSwitch(): Template {
  return build(
    "scenario-switch",
    "Scenario switch",
    "A selector (1, 2 or 3) picks Base, Upside or Downside for revenue growth, EBITDA margin, capex and tax into a Live column, and a three-year projection to unlevered free cash flow (no working capital) reads only that column.",
    [
      pad([title("Scenario switch")], COLUMNS, "title"),
      ...selectorRows(),
      blankRow(COLUMNS),
      header(["Assumptions", ...SCENARIOS, "Live"]),
      assumptionRow("Revenue growth (%)", ROW.growth, DEMO.growth),
      assumptionRow("EBITDA margin (%)", ROW.margin, DEMO.margin),
      assumptionRow("Capex (% of revenue)", ROW.capex, DEMO.capex),
      assumptionRow("Tax rate (%)", ROW.tax, DEMO.tax),
      blankRow(COLUMNS),
      header(["Projection (EUR m)", "Base year", "Year 1", "Year 2", "Year 3"]),
      ...projection(),
    ],
  );
}
