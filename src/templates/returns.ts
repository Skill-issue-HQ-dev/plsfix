// The returns block: entry equity, an exit in year 5 off EBITDA, a multiple and net
// debt into exit equity, MOIC and IRR over a cash-flow row, and MOIC at the exit
// multiple one step either side. Invariant: the cash-flow row has its sign change
// by construction (equity out in year 0, equity back in year 5), so IRR sees both.

import {
  blankRow,
  build,
  calc,
  formulaCell,
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
import { at, span } from "./refs";

const YEARS = 5;
// A label column, then the cash flows of year 0 to year 5.
const COLUMNS = YEARS + 2;
// The assumptions and answers sit in column 1; year n's cash flow in column n + 1.
const VALUE = 1;
const FIRST_FLOW = 1;

// Where each line sits, so a formula reads by what it names.
const ROW = {
  entry: 1,
  ebitda: 2,
  multiple: 3,
  ev: 4,
  netDebt: 5,
  exitEquity: 6,
  year: 8,
  flow: 9,
  moic: 11,
  irr: 12,
  step: 14,
  exitMultiples: 15,
} as const;

// EUR millions: 225.0 of equity in a 4.8x levered buyout of a 62.5 EBITDA business
// at 8.0x, sold at 8.0x on 92.0 of EBITDA with net debt paid down to 175.0.
const ENTRY_EQUITY = 225;
const EXIT_EBITDA = 92;
const EXIT_MULTIPLE = 8;
const NET_DEBT_AT_EXIT = 175;
const MULTIPLE_STEP = 1;
// Excel's own first guess for IRR.
const DEFAULT_GUESS = 0.1;

const value = (row: number): string => at(row, VALUE);

// The cash-flow row from year `from` to year `to` (0 to 5), as a range.
const flowRange = (from: number, to: number): string =>
  span(ROW.flow, FIRST_FLOW + from, ROW.flow, FIRST_FLOW + to);

function line(name: string, cell: TemplateCell): TemplateCell[] {
  return pad([label(name), cell], COLUMNS);
}

// The equity goes out in year 0, nothing moves in years 1 to 4 and the exit
// equity comes back in year 5. The middle years are typed zeros, never blanks:
// IRR skips a blank cell, which would squeeze the five years into two.
function cashFlows(): TemplateCell[] {
  const middle = Array.from({ length: YEARS - 1 }, () => input(0, "decimal"));
  return [
    label("Equity cash flow (EUR m)"),
    calc(`=-${value(ROW.entry)}`, "decimal"),
    ...middle,
    calc(`=${value(ROW.exitEquity)}`, "decimal"),
  ];
}

// The block's one IFERROR, and only here. IRR answers #NUM! when its cash flows
// never change sign (a typed negative entry equity, an exit that leaves no equity
// or an empty block) and when its search finds no rate; "n/a" says so in words.
// Two things keep a rate from being wrong instead: IRR skips a blank or a word
// and would squeeze the timeline, so all six flows must be numbers, and Excel's
// search is a short Newton walk that dies on a near-total loss from its own 10 %
// guess, so it starts at the rate a row with nothing in between has, MOIC^(1/5) - 1
// (10 % when the MOIC is not positive). Nothing else in the block is guarded, so
// every other error stays visible.
function irrCell(): TemplateCell {
  const flows = flowRange(0, YEARS);
  const moic = value(ROW.moic);
  const guess = `IF(N(${moic})>0,${moic}^(1/${YEARS})-1,${DEFAULT_GUESS})`;
  const squeezed = `COUNT(${flows})<${YEARS + 1}`;
  const rate = `IF(${squeezed},"n/a",IRR(${flows},${guess}))`;
  return total(`=IFERROR(${rate},"n/a")`, "percent");
}

// MOIC is every euro back over the equity in: years 1 to 5 of the cash-flow row,
// so a dividend counts beside the exit and MOIC and IRR read the same cash.
function moicCell(): TemplateCell {
  const back = flowRange(1, YEARS);
  return total(`=SUM(${back})/${value(ROW.entry)}`, "multiple");
}

// MOIC at the exit multiple one step below, at and one step above the base: the
// same EBITDA, net debt and interim cash, only the multiple moves, so the middle
// column is the MOIC above it.
function sensitivity(): TemplateCell[][] {
  const base = value(ROW.multiple);
  const step = value(ROW.step);
  const multiples = [`=${base}-${step}`, `=${base}`, `=${base}+${step}`];
  const moic = (column: number): string => {
    const interim = `SUM(${flowRange(1, YEARS - 1)})`;
    const ebitda = value(ROW.ebitda);
    const exit = `${ebitda}*${at(ROW.exitMultiples, column)}`;
    return `=(${interim}+${exit}-${value(ROW.netDebt)})/${value(ROW.entry)}`;
  };
  return [
    line("Multiple step", input(MULTIPLE_STEP, "multiple")),
    pad(
      [
        head("Exit multiple"),
        ...multiples.map((formula) =>
          formulaCell(formula, "header", "multiple"),
        ),
      ],
      COLUMNS,
      "header",
    ),
    pad(
      [
        label("MOIC at that multiple"),
        ...multiples.map((_formula, index) =>
          calc(moic(index + 1), "multiple"),
        ),
      ],
      COLUMNS,
    ),
  ];
}

export function returns(): Template {
  // Words, not numbers: the number formats show a typed 0 as a dash, and year 0
  // is the one header that must read "0".
  const years = Array.from({ length: YEARS + 1 }, (_unused, year) =>
    head(`Year ${String(year)}`),
  );
  return build(
    "returns",
    "Returns (MOIC / IRR)",
    "Entry equity, an exit in year 5 off EBITDA, an exit multiple and net debt into exit equity, MOIC and IRR over a cash-flow row, and MOIC at the exit multiple one step below and above.",
    [
      pad([title("Returns (MOIC / IRR)")], COLUMNS, "title"),
      line("Entry equity (EUR m)", input(ENTRY_EQUITY, "decimal")),
      line("Exit EBITDA (EUR m)", input(EXIT_EBITDA, "decimal")),
      line("Exit multiple", input(EXIT_MULTIPLE, "multiple")),
      line(
        "Exit EV (EUR m)",
        calc(`=${value(ROW.ebitda)}*${value(ROW.multiple)}`, "decimal"),
      ),
      line("Net debt at exit (EUR m)", input(NET_DEBT_AT_EXIT, "decimal")),
      line(
        "Exit equity (EUR m)",
        calc(`=${value(ROW.ev)}-${value(ROW.netDebt)}`, "decimal"),
      ),
      blankRow(COLUMNS),
      pad([head("EUR m"), ...years], COLUMNS, "header"),
      cashFlows(),
      blankRow(COLUMNS),
      pad([result("MOIC"), moicCell()], COLUMNS),
      pad([result("IRR"), irrCell()], COLUMNS),
      blankRow(COLUMNS),
      ...sensitivity(),
    ],
  );
}
