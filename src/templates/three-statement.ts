// The three-statement mini block: three years of an income statement, a balance
// sheet that opens on typed balances and a cash flow, in EUR millions, cash being
// the plug and a check reading zero while the sheet balances. Owns the grid and
// its demo case; the invariant is that interest reads the OPENING net debt.

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
  type TemplateCell,
  type TemplateFormat,
  title,
  total,
} from "../template-cells";
import { at } from "./refs";
import { yearColumns, yearHeader } from "./years";

const YEARS = 3;
// The balance sheet opens in the column before year one; the labels sit before it.
const LEAD = ["Opening"];
const OPENING_COLUMN = 1;
const COLUMNS = 1 + LEAD.length + YEARS;
const YEAR_COLUMNS = yearColumns(YEARS, LEAD.length);
const BALANCE_COLUMNS = [OPENING_COLUMN, ...YEAR_COLUMNS];
// The check is read to six places of EUR millions, which also wipes out the
// floating-point dust two equal sums can differ by.
const CHECK_DIGITS = 6;

// Where each line sits: the income statement, the balance sheet, the cash flow.
const ROW = {
  revenue: 2,
  margin: 3,
  ebitda: 4,
  da: 5,
  ebit: 6,
  rate: 7,
  interest: 8,
  ebt: 9,
  taxRate: 10,
  tax: 11,
  netIncome: 12,
  ppe: 14,
  wc: 15,
  cash: 16,
  debt: 17,
  equity: 18,
  netDebt: 19,
  check: 20,
  plusDa: 22,
  capex: 23,
  wcIncrease: 24,
  fcf: 25,
} as const;

// The demo case, EUR millions: a business growing 10 % a year, its margin up a
// point a year, 20 % tax, 7.5 % on 40 of opening net debt (50 of debt less 10 of
// cash) and a sheet of 60 PP&E and 20 working capital funded by 50 debt, 40 equity.
const DEMO = {
  revenue: [120, 132, 144],
  margin: [0.2, 0.21, 0.22],
  da: [6, 6.5, 7],
  rate: [0.075, 0.075, 0.075],
  taxRate: [0.2, 0.2, 0.2],
  capex: [8, 9, 9.5],
  wcIncrease: [2, 2.5, 2.5],
  ppe: 60,
  wc: 20,
  cash: 10,
  debt: 50,
  equity: 40,
};

// A line typed in for each of the three years; the Opening column holds nothing.
function typedRow(
  name: string,
  values: number[],
  format: TemplateFormat = "decimal",
): TemplateCell[] {
  return [label(name), blank(), ...values.map((value) => input(value, format))];
}

// A line of formulas, one per year, each built from its own year's column.
function derived(
  name: string,
  formula: (column: number) => string,
): TemplateCell[] {
  return [
    label(name),
    blank(),
    ...YEAR_COLUMNS.map((column) => calc(formula(column), "decimal")),
  ];
}

// The same, set apart as a subtotal.
function subtotal(
  name: string,
  formula: (column: number) => string,
): TemplateCell[] {
  return [
    result(name),
    blank("result"),
    ...YEAR_COLUMNS.map((column) => total(formula(column), "decimal")),
  ];
}

// Tax is charged on a profit only: a loss pays none and flows whole into net
// income, so no year books a refund the model has no way to claim. Interest is
// the typed rate on the net debt the year OPENS on, the column before its own, so
// it never reads anything the year itself computes.
function incomeStatement(): TemplateCell[][] {
  return [
    typedRow("Revenue (EUR m)", DEMO.revenue),
    typedRow("EBITDA margin (%)", DEMO.margin, "percent"),
    derived(
      "EBITDA (EUR m)",
      (c) => `=${at(ROW.revenue, c)}*${at(ROW.margin, c)}`,
    ),
    typedRow("D&A (EUR m)", DEMO.da),
    derived("EBIT (EUR m)", (c) => `=${at(ROW.ebitda, c)}-${at(ROW.da, c)}`),
    typedRow("Interest rate (% of opening net debt)", DEMO.rate, "percent"),
    derived(
      "Interest expense (EUR m, on opening net debt)",
      (c) => `=${at(ROW.rate, c)}*${at(ROW.netDebt, c - 1)}`,
    ),
    derived(
      "EBT, profit before tax (EUR m)",
      (c) => `=${at(ROW.ebit, c)}-${at(ROW.interest, c)}`,
    ),
    typedRow("Tax rate (%)", DEMO.taxRate, "percent"),
    derived(
      "Tax (EUR m, none on a loss)",
      (c) => `=MAX(${at(ROW.ebt, c)},0)*${at(ROW.taxRate, c)}`,
    ),
    subtotal(
      "Net income (EUR m)",
      (c) => `=${at(ROW.ebt, c)}-${at(ROW.tax, c)}`,
    ),
  ];
}

// A balance typed in the Opening column, then each year the column before it
// moved by that year's own flow.
function rolled(
  name: string,
  opening: number,
  move: (column: number) => string,
): TemplateCell[] {
  return [
    label(name),
    input(opening, "decimal"),
    ...YEAR_COLUMNS.map((column) => calc(move(column), "decimal")),
  ];
}

// Every flow moves its own balance and cash moves by the free cash flow alone:
// cash is the plug, so the sheet balances only while the flows add up.
function balanceSheet(): TemplateCell[][] {
  const before = (row: number, c: number): string => at(row, c - 1);
  return [
    rolled(
      "PP&E (EUR m, plus capex less D&A)",
      DEMO.ppe,
      (c) => `=${before(ROW.ppe, c)}+${at(ROW.capex, c)}-${at(ROW.da, c)}`,
    ),
    rolled(
      "Working capital (EUR m, plus the increase)",
      DEMO.wc,
      (c) => `=${before(ROW.wc, c)}+${at(ROW.wcIncrease, c)}`,
    ),
    rolled(
      "Cash (EUR m, plus free cash flow: the plug)",
      DEMO.cash,
      (c) => `=${before(ROW.cash, c)}+${at(ROW.fcf, c)}`,
    ),
    rolled(
      "Debt (EUR m, held flat)",
      DEMO.debt,
      (c) => `=${before(ROW.debt, c)}`,
    ),
    rolled(
      "Equity (EUR m, plus net income)",
      DEMO.equity,
      (c) => `=${before(ROW.equity, c)}+${at(ROW.netIncome, c)}`,
    ),
    netDebtRow(),
    checkRow(),
  ];
}

// A memo: debt less cash, the figure interest reads the year after.
function netDebtRow(): TemplateCell[] {
  return [
    label("Net debt (EUR m, debt less cash)"),
    ...BALANCE_COLUMNS.map((column) =>
      calc(`=${at(ROW.debt, column)}-${at(ROW.cash, column)}`, "decimal"),
    ),
  ];
}

// Zero while the sheet balances AND the net-debt memo is still debt less cash, in
// the Opening column and every year. Each gap is taken at its distance, so two
// gaps of opposite sign cannot cancel into a zero.
function checkRow(): TemplateCell[] {
  return [
    label("Check: assets less debt and equity, net debt tie (0)"),
    ...BALANCE_COLUMNS.map((column) => {
      const cell = (row: number): string => at(row, column);
      const assets = `${cell(ROW.ppe)}+${cell(ROW.wc)}+${cell(ROW.cash)}`;
      const sheet = `${assets}-${cell(ROW.debt)}-${cell(ROW.equity)}`;
      const memo = `${cell(ROW.netDebt)}-(${cell(ROW.debt)}-${cell(ROW.cash)})`;
      return calc(
        `=ROUND(ABS(${sheet})+ABS(${memo}),${CHECK_DIGITS})`,
        "decimal",
      );
    }),
  ];
}

// The cash flow starts from net income: add D&A back, take the capex and the
// working capital the business absorbed.
function cashFlow(): TemplateCell[][] {
  return [
    derived("Plus: D&A (EUR m)", (c) => `=${at(ROW.da, c)}`),
    typedRow("Less: capex (EUR m)", DEMO.capex),
    typedRow("Less: increase in working capital (EUR m)", DEMO.wcIncrease),
    subtotal("Free cash flow (EUR m)", (c) => {
      const earned = `${at(ROW.netIncome, c)}+${at(ROW.plusDa, c)}`;
      return `=${earned}-${at(ROW.capex, c)}-${at(ROW.wcIncrease, c)}`;
    }),
  ];
}

export function threeStatement(): Template {
  return build(
    "three-statement",
    "Three-statement mini",
    "An income statement, a balance sheet that opens on typed balances and a cash flow over three years: interest on the opening net debt, cash as the plug, and a check that it balances each year.",
    [
      pad([title("Three-statement mini")], COLUMNS, "title"),
      yearHeader(YEARS, LEAD),
      ...incomeStatement(),
      blankRow(COLUMNS),
      ...balanceSheet(),
      blankRow(COLUMNS),
      ...cashFlow(),
    ],
  );
}
