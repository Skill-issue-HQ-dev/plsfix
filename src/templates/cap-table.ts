// The cap table block: pre-money, new money and a new option pool into post-money,
// price per share and new shares, with six holders before and after the round, in
// millions of EUR and of shares. Invariant: the pool sits inside the pre-money, so
// the new money buys exactly its post-money stake and each side adds up to 100 %.

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
  type TemplateFormat,
  title,
  total,
} from "../template-cells";
import { at, span } from "./refs";

const COLUMNS = 6;
const COL = {
  holder: 0,
  before: 1,
  beforePct: 2,
  issued: 3,
  after: 4,
  afterPct: 5,
} as const;
// The round's own lines keep their value in the "before" column.
const VALUE = 1;
// A percent column adds up to 100 %, and the new money's stake matches what it
// bought, within this many percentage points as a fraction.
const PERCENT_TOLERANCE = 0.0001;

// EUR millions and millions of shares: a 40.0 pre-money round of 10.0 with a 10 %
// pool, over 10.0 shares held by five holders; price 3.50, a 20 % stake.
const PRE_MONEY = 40;
const NEW_MONEY = 10;
const NEW_POOL = 0.1;
const POOL_HOLDER = "Option pool";
const HOLDERS: [name: string, shares: number][] = [
  ["Founders", 5.5],
  ["Series A", 2],
  ["Series B", 1],
  [POOL_HOLDER, 0.5],
  ["Other holders", 1],
];

// Rows, top to bottom: the round, a gap, the table's header, the five holders
// and the new investors under them, the total, a gap and the two checks.
const ROW = {
  pre: 1,
  money: 2,
  post: 3,
  stake: 4,
  pool: 5,
  before: 6,
  after: 7,
  price: 8,
  issued: 9,
  poolShares: 10,
  firstHolder: 13,
} as const;
const NEW_INVESTORS_ROW = ROW.firstHolder + HOLDERS.length;
const TOTAL_ROW = NEW_INVESTORS_ROW + 1;

const value = (row: number): string => at(row, VALUE);

// Post-money is pre-money plus new money, and the new money's stake is its share
// of that. The pool is a share of the shares AFTER the round and part of the
// pre-money, so the old holders carry it: they keep what is left once the new
// money and the pool are taken out of the whole, which gives the share count
// after the round, and the price follows from post-money over it.
function round(): TemplateCell[][] {
  const stake = value(ROW.stake);
  const lines: [name: string, cell: TemplateCell][] = [
    ["Pre-money valuation (EUR m)", input(PRE_MONEY, "decimal")],
    ["New money (EUR m)", input(NEW_MONEY, "decimal")],
    [
      "Post-money valuation (EUR m)",
      calc(`=${value(ROW.pre)}+${value(ROW.money)}`, "decimal"),
    ],
    [
      "New money as % of post-money",
      calc(`=${value(ROW.money)}/${value(ROW.post)}`, "percent"),
    ],
    ["New option pool (% of post-money)", input(NEW_POOL, "percent")],
    [
      "Shares before the round (m)",
      calc(`=${at(TOTAL_ROW, COL.before)}`, "decimal"),
    ],
    [
      "Shares after the round (m)",
      calc(`=${value(ROW.before)}/(1-${stake}-${value(ROW.pool)})`, "decimal"),
    ],
    [
      "Price per share (EUR)",
      calc(`=${value(ROW.post)}/${value(ROW.after)}`, "decimal2"),
    ],
    ["New shares issued (m)", calc(`=${stake}*${value(ROW.after)}`, "decimal")],
    [
      "New option pool shares (m)",
      calc(`=${value(ROW.pool)}*${value(ROW.after)}`, "decimal"),
    ],
  ];
  return lines.map(([name, cell]) => pad([label(name), cell], COLUMNS));
}

// One holder: shares before, their stake, any new shares, shares after and the
// stake after. `fresh` is the round's cell the new shares come from, if any.
function holderRow(
  name: string,
  before: TemplateCell,
  row: number,
  fresh?: number,
): TemplateCell[] {
  return [
    label(name),
    before,
    calc(`=${at(row, COL.before)}/${at(TOTAL_ROW, COL.before)}`, "percent"),
    fresh === undefined ? blank() : calc(`=${value(fresh)}`, "decimal"),
    calc(`=${at(row, COL.before)}+${at(row, COL.issued)}`, "decimal"),
    calc(`=${at(row, COL.after)}/${at(TOTAL_ROW, COL.after)}`, "percent"),
  ];
}

function totalRow(): TemplateCell[] {
  const sum = (column: number, format: TemplateFormat): TemplateCell =>
    total(
      `=SUM(${span(ROW.firstHolder, column, NEW_INVESTORS_ROW, column)})`,
      format,
    );
  return [
    result("Total"),
    sum(COL.before, "decimal"),
    sum(COL.beforePct, "percent"),
    sum(COL.issued, "decimal"),
    sum(COL.after, "decimal"),
    sum(COL.afterPct, "percent"),
  ];
}

// Each side's stakes add to 100 %; and the new investors own what their money
// bought, over a share count that is positive (a pool and a stake that take
// more than all of the company make it negative).
function checks(): TemplateCell[][] {
  const tolerance = String(PERCENT_TOLERANCE);
  const near = (cell: string, target: string): string =>
    `ABS(${cell}-${target})<${tolerance}`;
  const sides = [
    near(at(TOTAL_ROW, COL.beforePct), "1"),
    near(at(TOTAL_ROW, COL.afterPct), "1"),
  ].join(",");
  const bought = near(at(NEW_INVESTORS_ROW, COL.afterPct), value(ROW.stake));
  return [
    pad(
      [
        label("Check: each side adds up to 100 %"),
        calc(`=IF(AND(${sides}),"OK","CHECK")`, "decimal"),
      ],
      COLUMNS,
    ),
    pad(
      [
        label("Check: new money buys its stake"),
        calc(
          `=IF(AND(${value(ROW.after)}>0,${bought}),"OK","CHECK")`,
          "decimal",
        ),
      ],
      COLUMNS,
    ),
  ];
}

export function capTable(): Template {
  const header = [
    "Holder",
    "Shares before (m)",
    "% before",
    "New shares (m)",
    "Shares after (m)",
    "% after",
  ].map((name) => head(name));
  return build(
    "cap-table",
    "Cap table",
    "A pre-money valuation, new money and a new option pool (a share of the post-money shares, inside the pre-money) into post-money, the price per share and the new shares, with six holders' shares and stakes before and after the round.",
    [
      pad([title("Cap table")], COLUMNS, "title"),
      ...round(),
      blankRow(COLUMNS),
      header,
      ...HOLDERS.map(([name, shares], index) =>
        holderRow(
          name,
          input(shares, "decimal"),
          ROW.firstHolder + index,
          name === POOL_HOLDER ? ROW.poolShares : undefined,
        ),
      ),
      holderRow("New investors", blank(), NEW_INVESTORS_ROW, ROW.issued),
      totalRow(),
      blankRow(COLUMNS),
      ...checks(),
    ],
  );
}
