// The trading-comps block: six peers' EV, revenue, EBITDA, net income and net
// debt into EV / Revenue, EV / EBITDA and P / E, their six statistics, and a
// target's implied EV off the median multiples, in EUR millions. Invariant: a
// multiple over a loss, a blank or a typed word is left empty, so no statistic
// or implied EV ever includes it, while a genuine error still shows.

import {
  blank,
  blankRow,
  build,
  calc,
  head,
  input,
  inputText,
  label,
  pad,
  result,
  type Template,
  type TemplateCell,
  title,
  total,
} from "../template-cells";
import { at, span } from "./refs";
import { STAT_COUNT, STAT_ROW, statRows } from "./stat-rows";

const COLUMNS = 9;
const COL = {
  name: 0,
  ev: 1,
  revenue: 2,
  ebitda: 3,
  income: 4,
  debt: 5,
  evRevenue: 6,
  evEbitda: 7,
  pe: 8,
} as const;
const MULTIPLE_COLUMNS = [COL.evRevenue, COL.evEbitda, COL.pe];
// The column the implied EVs, their range and the midpoint sit in.
const IMPLIED_VALUE = 3;

// A peer as typed into the block: name, EV, revenue, EBITDA, net income and net
// debt, EUR millions.
type Peer = [
  name: string,
  ev: number,
  revenue: number,
  ebitda: number,
  income: number,
  debt: number,
];

// Six mid-caps: EV 1.1x to 1.6x revenue, 8.7x to 9.9x EBITDA, 15x to 19x
// earnings, the spread a sector trades in.
const PEERS: Peer[] = [
  ["Peer A", 1240, 960, 138, 62, 310],
  ["Peer B", 860, 710, 91, 41, 190],
  ["Peer C", 2150, 1380, 248, 115, 420],
  ["Peer D", 540, 480, 59, 22, 145],
  ["Peer E", 1680, 1150, 176, 78, 390],
  ["Peer F", 730, 520, 74, 30, 160],
];
const TARGET = {
  name: "Target Co",
  revenue: 410,
  ebitda: 52,
  income: 21,
  debt: 120,
};

// Rows, top to bottom: title, header, the peers, six statistics, a gap, the
// target's header and row, a gap, the implied table, its range and midpoint.
const FIRST_PEER = 2;
const LAST_PEER = FIRST_PEER + PEERS.length - 1;
const STAT_FIRST = LAST_PEER + 1;
const MEDIAN_ROW = STAT_FIRST + STAT_ROW.median;
const TARGET_ROW = STAT_FIRST + STAT_COUNT + 2;
const IMPLIED_FIRST = TARGET_ROW + 3;
const IMPLIED_LAST = IMPLIED_FIRST + 2;

const HEADER = [
  "Company (EUR m)",
  "EV",
  "Revenue",
  "EBITDA",
  "Net income",
  "Net debt",
  "EV / Revenue",
  "EV / EBITDA",
  "P / E",
];

// A multiple over a zero, a loss, a blank or a typed word is not meaningful: it
// stays empty so the statistics skip it, which is what NM means on a comps page.
// N() reads a word or a blank as zero and lets a genuine error through, so an
// "n.a." leaves the row empty while a real #N/A still shows. P / E takes equity
// value, EV less net debt, over net income, and is empty when that is not positive.
function peerRow(peer: Peer, row: number): TemplateCell[] {
  const cell = (column: number): string => at(row, column);
  const [ev, revenue, ebitda, income, debt] = [
    COL.ev,
    COL.revenue,
    COL.ebitda,
    COL.income,
    COL.debt,
  ].map(cell) as [string, string, string, string, string];
  const positive = (reference: string): string => `N(${reference})>0`;
  const multiple = (
    tests: string[],
    top: string,
    bottom: string,
  ): TemplateCell =>
    calc(`=IF(AND(${tests.join(",")}),${top}/${bottom},"")`, "multiple");
  const [name, ...figures] = peer;
  return [
    inputText(name),
    ...figures.map((figure) => input(figure, "decimal")),
    multiple([positive(ev), positive(revenue)], ev, revenue),
    multiple([positive(ev), positive(ebitda)], ev, ebitda),
    multiple(
      [positive(ev), positive(income), `N(${ev})-N(${debt})>0`],
      `(${ev}-N(${debt}))`,
      income,
    ),
  ];
}

function targetRows(): TemplateCell[][] {
  return [
    pad(
      [
        head("Target"),
        blank("header"),
        head("Revenue"),
        head("EBITDA"),
        head("Net income"),
        head("Net debt"),
      ],
      COLUMNS,
      "header",
    ),
    pad(
      [
        inputText(TARGET.name),
        blank(),
        input(TARGET.revenue, "decimal"),
        input(TARGET.ebitda, "decimal"),
        input(TARGET.income, "decimal"),
        input(TARGET.debt, "decimal"),
      ],
      COLUMNS,
    ),
  ];
}

// One method's line: the target's own metric, the median multiple read off the
// statistics, and their product. A method that means nothing (a metric that is
// zero, negative or a word, a multiple left empty) drops out as "". P / E prices
// equity, so the target's net debt is added back to land on an enterprise value
// and has to be readable too.
function impliedRow(
  name: string,
  row: number,
  metric: number,
  multiple: number,
  note?: string,
): TemplateCell[] {
  const [value, factor] = [at(row, 1), at(row, 2)];
  const debt = at(TARGET_ROW, COL.debt);
  const priced = note === undefined ? "" : `+${debt}`;
  const unreadable = [`N(${value})<=0`, `ISTEXT(${factor})`];
  if (note !== undefined) unreadable.push(`ISTEXT(${debt})`);
  const cells = [
    label(name),
    calc(`=${at(TARGET_ROW, metric)}`, "decimal"),
    calc(`=${at(MEDIAN_ROW, multiple)}`, "multiple"),
    calc(
      `=IF(OR(${unreadable.join(",")}),"",${value}*${factor}${priced})`,
      "decimal",
    ),
  ];
  return pad(note === undefined ? cells : [...cells, label(note)], COLUMNS);
}

function impliedRows(): TemplateCell[][] {
  const column = span(
    IMPLIED_FIRST,
    IMPLIED_VALUE,
    IMPLIED_LAST,
    IMPLIED_VALUE,
  );
  const low = IMPLIED_LAST + 1;
  const high = low + 1;
  const range = (name: string, fn: string): TemplateCell[] =>
    pad(
      [
        label(name),
        blank(),
        blank(),
        calc(`=IF(COUNT(${column})=0,"",${fn}(${column}))`, "decimal"),
      ],
      COLUMNS,
    );
  const bounds = span(low, IMPLIED_VALUE, high, IMPLIED_VALUE);
  return [
    pad(
      [
        head("Implied EV (EUR m)"),
        head("Target metric"),
        head("Median multiple"),
        head("Implied EV"),
      ],
      COLUMNS,
      "header",
    ),
    impliedRow("EV / Revenue", IMPLIED_FIRST, COL.revenue, COL.evRevenue),
    impliedRow("EV / EBITDA", IMPLIED_FIRST + 1, COL.ebitda, COL.evEbitda),
    impliedRow("P / E", IMPLIED_LAST, COL.income, COL.pe, "plus net debt"),
    range("Implied EV, low", "MIN"),
    range("Implied EV, high", "MAX"),
    pad(
      [
        result("Implied EV, midpoint"),
        blank("result"),
        blank("result"),
        total(`=IF(COUNT(${bounds})=0,"",AVERAGE(${bounds}))`, "decimal"),
      ],
      COLUMNS,
      "result",
    ),
  ];
}

export function tradingComps(): Template {
  return build(
    "trading-comps",
    "Trading comps",
    "Six peers' EV, revenue, EBITDA, net income and net debt into EV / Revenue, EV / EBITDA and P / E, six statistics per multiple, and a target's implied EV from the median multiples.",
    [
      pad([title("Trading comps")], COLUMNS, "title"),
      HEADER.map((name) => head(name)),
      ...PEERS.map((peer, index) => peerRow(peer, FIRST_PEER + index)),
      ...statRows({
        first: FIRST_PEER,
        last: LAST_PEER,
        columns: MULTIPLE_COLUMNS,
        width: COLUMNS,
      }),
      blankRow(COLUMNS),
      ...targetRows(),
      blankRow(COLUMNS),
      ...impliedRows(),
    ],
  );
}
