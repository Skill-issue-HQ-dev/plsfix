// The revenue-build block: a year-one volume and price, each grown by its own
// typed yearly rate, into revenue in EUR millions, its yearly growth and the
// year 1 to 5 CAGR. Owns the grid and its demo case; the invariant is that every
// placeholder names a cell inside the 11 by 6 block and no year reads a later one.

import {
  blank,
  blankRow,
  build,
  calc,
  input,
  label,
  pad,
  result,
  type Template,
  type TemplateFormat,
  type TemplateCell,
  title,
  total,
} from "../template-cells";
import { at } from "./refs";
import { yearColumns, yearHeader } from "./years";

const YEARS = 5;
const COLUMNS = YEARS + 1;
const YEAR_COLUMNS = yearColumns(YEARS);
const FIRST_YEAR = 1;
const LAST_YEAR = YEARS;
const LATER_YEARS = YEAR_COLUMNS.slice(1);
// A volume in thousand units times a price in EUR is EUR thousand; a thousand of
// those make one EUR million.
const THOUSANDS_PER_MILLION = 1_000;
// Years between the first revenue and the last: the root the CAGR takes.
const INTERVALS = YEARS - 1;

// Where each line sits: a level, the rate that grows it, then the revenue.
const ROW = {
  header: 1,
  volume: 2,
  volumeGrowth: 3,
  price: 4,
  priceGrowth: 5,
  revenue: 7,
  revenueGrowth: 8,
  cagr: 10,
} as const;

// The demo case: the year-one level, then the typed rate into years two to five.
const VOLUME = { level: 1200, growth: [0.05, 0.05, 0.04, 0.04] };
const PRICE = { level: 25, growth: [0.03, 0.03, 0.025, 0.025] };

// A level typed in year one, then each year the prior year grown by the rate
// typed on the row beneath it, all shown in the same format.
function levelRow(
  name: string,
  row: number,
  level: number,
  format: TemplateFormat,
): TemplateCell[] {
  return [
    label(name),
    input(level, format),
    ...LATER_YEARS.map((column) =>
      calc(`=${at(row, column - 1)}*(1+${at(row + 1, column)})`, format),
    ),
  ];
}

// The typed yearly rates; year one has no year before it to grow from.
function rateRow(name: string, rates: number[]): TemplateCell[] {
  return [label(name), blank(), ...rates.map((rate) => input(rate, "percent"))];
}

function revenueRow(): TemplateCell[] {
  return [
    result("Revenue (EUR m)"),
    ...YEAR_COLUMNS.map((column) =>
      total(
        `=${at(ROW.volume, column)}*${at(ROW.price, column)}/${THOUSANDS_PER_MILLION}`,
        "decimal",
      ),
    ),
  ];
}

// Growth over the year before; a year with no revenue behind it is not meaningful.
function growthRow(): TemplateCell[] {
  return [
    label("Revenue growth (%)"),
    blank(),
    ...LATER_YEARS.map((column) => {
      const before = at(ROW.revenue, column - 1);
      const now = at(ROW.revenue, column);
      return calc(`=IF(${before}>0,${now}/${before}-1,"n.m.")`, "percent");
    }),
  ];
}

// The compound rate from the first year's revenue to the last, by plain formula
// so a workbook without the add-in still computes it.
function cagrRow(): TemplateCell[] {
  const first = at(ROW.revenue, FIRST_YEAR);
  const last = at(ROW.revenue, LAST_YEAR);
  const rate =
    `=IF(AND(${first}>0,${last}>0),` +
    `(${last}/${first})^(1/${INTERVALS})-1,"n.m.")`;
  return pad(
    [result(`Revenue CAGR (year 1 to ${YEARS})`), total(rate, "percent")],
    COLUMNS,
    "result",
  );
}

export function revenueBuild(): Template {
  return build(
    "revenue-build",
    "Revenue build",
    "A year-one volume and price, each grown by its own typed yearly rate, into revenue in EUR millions, its yearly growth and the year 1 to 5 CAGR.",
    [
      pad([title("Revenue build")], COLUMNS, "title"),
      yearHeader(YEARS),
      levelRow("Volume (thousand units)", ROW.volume, VOLUME.level, "decimal"),
      rateRow("Volume growth (%)", VOLUME.growth),
      levelRow("Price (EUR per unit)", ROW.price, PRICE.level, "decimal2"),
      rateRow("Price growth (%)", PRICE.growth),
      blankRow(COLUMNS),
      revenueRow(),
      growthRow(),
      blankRow(COLUMNS),
      cagrRow(),
    ],
  );
}
