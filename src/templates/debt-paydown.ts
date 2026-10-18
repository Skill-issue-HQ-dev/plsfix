// The debt paydown block: a five-year senior term loan with interest on the
// opening balance, mandatory amortisation of the original loan, a cash sweep and
// leverage, in EUR millions. Invariant: no formula reads a closing balance of its
// own year, so the block has no circular reference and no balance goes below zero.

import {
  blankRow,
  build,
  calc,
  head,
  input,
  label,
  numberCell,
  pad,
  result,
  type Template,
  type TemplateCell,
  title,
  total,
} from "../template-cells";
import { at } from "./refs";

const YEARS = 5;
const COLUMNS = YEARS + 1;
// The four assumptions sit in column 1; the five years run over columns 1 to 5.
const VALUE = 1;

// Where each line sits, so a formula reads by what it names.
const ROW = {
  loan: 1,
  rate: 2,
  amortisation: 3,
  sweep: 4,
  cfads: 7,
  ebitda: 8,
  opening: 10,
  interest: 11,
  mandatory: 12,
  available: 13,
  swept: 14,
  closing: 15,
} as const;

// EUR millions: a 4.0x senior loan on 62.5 of EBITDA that grows to 92.0, cash
// flow for debt service of about 60 % of EBITDA, half of what is left swept.
const LOAN = 250;
const CFADS = [40, 44, 48, 52, 56];
const EBITDA = [68, 74, 80, 86, 92];

const assumption = (row: number): string => at(row, VALUE);

function line(name: string, cell: TemplateCell): TemplateCell[] {
  return pad([label(name), cell], COLUMNS);
}

const typedYears = (name: string, values: number[]): TemplateCell[] => [
  label(name),
  ...values.map((value) => input(value, "decimal")),
];

// `make` is given the year's column, 1 to 5.
const formulaYears = (
  name: TemplateCell,
  make: (column: number) => TemplateCell,
): TemplateCell[] => [
  name,
  ...Array.from({ length: YEARS }, (_unused, index) => make(index + 1)),
];

// The first year opens on the loan at close, every later one on the balance the
// year before closed on.
function opening(column: number): TemplateCell {
  const from =
    column === 1 ? assumption(ROW.loan) : at(ROW.closing, column - 1);
  return calc(`=${from}`, "decimal");
}

// Interest comes off the opening balance, never an average: an average would
// read the closing balance, the closing balance reads the sweep, and the sweep
// reads the interest.
function interest(column: number): TemplateCell {
  return calc(`=${at(ROW.opening, column)}*${assumption(ROW.rate)}`, "decimal");
}

// A fixed slice of the ORIGINAL loan, never more than is still owed.
function mandatory(column: number): TemplateCell {
  const slice = `${assumption(ROW.amortisation)}*${assumption(ROW.loan)}`;
  return calc(`=MIN(${slice},${at(ROW.opening, column)})`, "decimal");
}

// The year's cash for debt service less its interest and its mandatory slice. A
// shortfall shows as a negative number; it is not hidden and not borrowed.
function available(column: number): TemplateCell {
  const cash = at(ROW.cfads, column);
  const interest = at(ROW.interest, column);
  const mandatory = at(ROW.mandatory, column);
  return calc(`=${cash}-${interest}-${mandatory}`, "decimal");
}

// The sweep share of the cash left (none when there is none), never more than the
// balance still owed once the mandatory slice is paid.
function swept(column: number): TemplateCell {
  const left = `MAX(0,${at(ROW.available, column)})`;
  const owed = `${at(ROW.opening, column)}-${at(ROW.mandatory, column)}`;
  return calc(`=MIN(${assumption(ROW.sweep)}*${left},${owed})`, "decimal");
}

function closing(column: number): TemplateCell {
  const opening = at(ROW.opening, column);
  const mandatory = at(ROW.mandatory, column);
  return total(`=${opening}-${mandatory}-${at(ROW.swept, column)}`, "decimal");
}

// A multiple over a zero, negative or typed-word EBITDA is not meaningful: it says
// n.m. N() makes a word a 0, because text ranks above every number in a comparison
// and "n.a." > 0 would let the guard divide by the word.
function leverage(column: number): TemplateCell {
  const ebitda = at(ROW.ebitda, column);
  return calc(
    `=IF(N(${ebitda})>0,${at(ROW.closing, column)}/${ebitda},"n.m.")`,
    "multiple",
  );
}

export function debtPaydown(): Template {
  const years = Array.from({ length: YEARS }, (_unused, index) =>
    numberCell(index + 1, "header", "whole"),
  );
  return build(
    "debt-paydown",
    "Debt paydown (cash sweep)",
    "A five-year senior term loan: interest on the opening balance, mandatory amortisation of the original loan and a cash sweep out of the cash flow available for debt service, with the closing balance and leverage per year.",
    [
      pad([title("Debt paydown (cash sweep)")], COLUMNS, "title"),
      line("Term loan at close (EUR m)", input(LOAN, "decimal")),
      line("Interest rate", input(0.07, "percent")),
      line("Mandatory amortisation (% of original)", input(0.05, "percent")),
      line("Cash sweep (% of excess cash)", input(0.5, "percent")),
      blankRow(COLUMNS),
      [head("Year (EUR m)"), ...years],
      typedYears("CFADS", CFADS),
      typedYears("EBITDA", EBITDA),
      blankRow(COLUMNS),
      formulaYears(label("Opening balance"), opening),
      formulaYears(label("Interest"), interest),
      formulaYears(label("Mandatory amortisation"), mandatory),
      formulaYears(label("Cash available for sweep"), available),
      formulaYears(label("Cash sweep"), swept),
      formulaYears(result("Closing balance"), closing),
      blankRow(COLUMNS),
      formulaYears(label("Leverage (closing debt / EBITDA)"), leverage),
    ],
  );
}
