// The capex-and-depreciation block: revenue and capex over five years and a PP&E
// roll-forward in gross and net, straight-line on the gross balance with half a
// year on each year's capex, held at the net book value there is to write off.
// Owns the grid and its demo case; a year reads only itself and the year before.

import {
  blankRow,
  build,
  calc,
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
import { yearColumns, yearHeader } from "./years";

const YEARS = 5;
const COLUMNS = YEARS + 1;
const YEAR_COLUMNS = yearColumns(YEARS);
const LATER_YEARS = YEAR_COLUMNS.slice(1);
// Capex goes into service on average mid-year, so the year it is spent it earns
// half a year's depreciation.
const HALF_YEAR = 2;
// The one typed life sits in the first number column, under the labels.
const LIFE_COLUMN = 1;

// Where each line sits: the life, the year strip, revenue and capex, the gross
// roll, then the net roll with its depreciation.
const ROW = {
  life: 1,
  revenue: 4,
  capex: 5,
  capexShare: 6,
  grossOpening: 8,
  grossClosing: 9,
  netOpening: 11,
  plusCapex: 12,
  depreciation: 13,
  netClosing: 14,
} as const;

// The demo case, EUR millions: a ten-year life over 200 of PP&E, gross and net
// alike, and capex near a tenth of revenue.
const DEMO = {
  life: 10,
  gross: 200,
  net: 200,
  revenue: [300, 320, 340, 360, 380],
  capex: [30, 33, 36, 38, 40],
};

const LIFE = at(ROW.life, LIFE_COLUMN);

// A line typed in for each of the five years.
function typedRow(name: string, values: number[]): TemplateCell[] {
  return [label(name), ...values.map((value) => input(value, "decimal"))];
}

// Capex over revenue; a year with no revenue, or text where the revenue should be,
// has no share worth showing. N() reads text as 0, so "n.a." is not > 0.
function shareRow(): TemplateCell[] {
  return [
    label("Capex (% of revenue)"),
    ...YEAR_COLUMNS.map((column) => {
      const revenue = at(ROW.revenue, column);
      const capex = at(ROW.capex, column);
      return calc(`=IF(N(${revenue})>0,${capex}/${revenue},"n.m.")`, "percent");
    }),
  ];
}

// The balance a year opens on: typed for year one, then the year before's close.
function openingRow(
  name: string,
  typedValue: number,
  closing: number,
): TemplateCell[] {
  return [
    label(name),
    input(typedValue, "decimal"),
    ...LATER_YEARS.map((column) =>
      calc(`=${at(closing, column - 1)}`, "decimal"),
    ),
  ];
}

// Gross PP&E closes at the opening plus the year's capex.
function grossClosingRow(): TemplateCell[] {
  return [
    label("Closing gross PP&E (EUR m)"),
    ...YEAR_COLUMNS.map((column) => {
      const opening = at(ROW.grossOpening, column);
      const capex = at(ROW.capex, column);
      return calc(`=${opening}+${capex}`, "decimal");
    }),
  ];
}

function capexLine(): TemplateCell[] {
  return [
    label("Plus: capex (EUR m)"),
    ...YEAR_COLUMNS.map((column) =>
      calc(`=${at(ROW.capex, column)}`, "decimal"),
    ),
  ];
}

// Straight-line on gross: the opening gross over the life, plus the year's capex
// over the life for half a year, and never more than the net book value left to
// write off. A life of zero or none divides by zero on purpose and MIN passes the
// error on, so the schedule shows it instead of a depreciation of nothing.
function depreciationRow(): TemplateCell[] {
  return [
    label(
      "Less: depreciation (EUR m, straight-line on gross, half-year on new capex)",
    ),
    ...YEAR_COLUMNS.map((column) => {
      const net = at(ROW.netOpening, column);
      const gross = at(ROW.grossOpening, column);
      const capex = at(ROW.plusCapex, column);
      const level = `${gross}/${LIFE}+${capex}/${LIFE}/${HALF_YEAR}`;
      return calc(`=MIN(${net}+${capex},${level})`, "decimal");
    }),
  ];
}

function closingRow(): TemplateCell[] {
  return [
    result("Closing net PP&E (EUR m)"),
    ...YEAR_COLUMNS.map((column) => {
      const opening = at(ROW.netOpening, column);
      const capex = at(ROW.plusCapex, column);
      const depreciation = at(ROW.depreciation, column);
      return total(`=${opening}+${capex}-${depreciation}`, "decimal");
    }),
  ];
}

export function capexDepreciation(): Template {
  return build(
    "capex-depreciation",
    "Capex and depreciation",
    "Revenue and capex for five years, capex as a share of revenue, and a PP&E roll-forward: straight-line on gross PP&E, half a year on new capex, net PP&E never below zero.",
    [
      pad([title("Capex and depreciation")], COLUMNS, "title"),
      pad([label("Useful life (years)"), input(DEMO.life, "decimal")], COLUMNS),
      blankRow(COLUMNS),
      yearHeader(YEARS),
      typedRow("Revenue (EUR m)", DEMO.revenue),
      typedRow("Capex (EUR m)", DEMO.capex),
      shareRow(),
      blankRow(COLUMNS),
      openingRow("Opening gross PP&E (EUR m)", DEMO.gross, ROW.grossClosing),
      grossClosingRow(),
      blankRow(COLUMNS),
      openingRow("Opening net PP&E (EUR m)", DEMO.net, ROW.netClosing),
      capexLine(),
      depreciationRow(),
      closingRow(),
    ],
  );
}
