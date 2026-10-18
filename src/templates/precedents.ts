// The precedent-transactions block: six deals' EV and EBITDA into EV / EBITDA,
// its six statistics, and the target's implied EV at the 25th percentile, 75th
// percentile and median multiple, in EUR millions. Invariant: a multiple over a
// zero, negative or typed-word EBITDA is left empty, so no statistic or implied
// EV ever includes it, while a genuine error still shows.

import {
  blank,
  blankRow,
  build,
  calc,
  dateInput,
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
import { at } from "./refs";
import { STAT_COUNT, STAT_ROW, statRows } from "./stat-rows";

const COLUMNS = 6;
const COL = {
  date: 0,
  target: 1,
  acquirer: 2,
  ev: 3,
  ebitda: 4,
  multiple: 5,
} as const;

// A deal as typed into the block: announced year, month and day, target,
// acquirer, then EV and EBITDA in EUR millions.
type Deal = [
  year: number,
  month: number,
  day: number,
  target: string,
  acquirer: string,
  ev: number,
  ebitda: number,
];

// Six deals, newest first: 7.6x to 11.6x EBITDA, the spread a mid-market sector
// clears at.
const DEALS: Deal[] = [
  [2025, 11, 14, "Target A", "Acquirer A", 420, 44],
  [2025, 6, 30, "Target B", "Acquirer B", 310, 33],
  [2025, 2, 19, "Target C", "Acquirer C", 950, 82],
  [2024, 10, 8, "Target D", "Acquirer D", 275, 36],
  [2024, 5, 22, "Target E", "Acquirer E", 640, 61],
  [2023, 12, 12, "Target F", "Acquirer F", 385, 40],
];
const TARGET_EBITDA = 52;

// Rows, top to bottom: title, header, the deals, six statistics, a gap, the
// target's EBITDA and the three implied EVs.
const FIRST_DEAL = 2;
const LAST_DEAL = FIRST_DEAL + DEALS.length - 1;
const STAT_FIRST = LAST_DEAL + 1;
const EBITDA_ROW = STAT_FIRST + STAT_COUNT + 1;

const HEADER = [
  "Date",
  "Target",
  "Acquirer",
  "EV (EUR m)",
  "EBITDA (EUR m)",
  "EV / EBITDA",
];

// A multiple over a zero, a loss, a blank or a typed word is not meaningful: it
// stays empty so the statistics skip it. N() reads a word or a blank as zero and
// lets a genuine error through, so "n.a." leaves the row empty, a real #N/A shows.
function dealRow(deal: Deal, row: number): TemplateCell[] {
  const ev = at(row, COL.ev);
  const ebitda = at(row, COL.ebitda);
  const [year, month, day, target, acquirer, evValue, ebitdaValue] = deal;
  return [
    dateInput(year, month, day),
    inputText(target),
    inputText(acquirer),
    input(evValue, "decimal"),
    input(ebitdaValue, "decimal"),
    calc(`=IF(AND(N(${ev})>0,N(${ebitda})>0),${ev}/${ebitda},"")`, "multiple"),
  ];
}

// The target's EBITDA at one statistic's multiple, empty when either means
// nothing: an EBITDA that is zero, negative or a word, or a multiple left empty.
function impliedEv(statRowIndex: number): string {
  const metric = at(EBITDA_ROW, COL.ebitda);
  const multiple = at(statRowIndex, COL.multiple);
  return `=IF(OR(N(${metric})<=0,ISTEXT(${multiple})),"",${metric}*${multiple})`;
}

// The value sits in the EV column, where an enterprise value belongs.
function impliedRow(name: string, statRowIndex: number): TemplateCell[] {
  return pad(
    [label(name), blank(), blank(), calc(impliedEv(statRowIndex), "decimal")],
    COLUMNS,
  );
}

function impliedRows(): TemplateCell[][] {
  const at25 = STAT_FIRST + STAT_ROW.p25;
  const at75 = STAT_FIRST + STAT_ROW.p75;
  return [
    pad(
      [
        label("Target EBITDA"),
        blank(),
        blank(),
        blank(),
        input(TARGET_EBITDA, "decimal"),
      ],
      COLUMNS,
    ),
    impliedRow("Implied EV at the 25th percentile multiple", at25),
    impliedRow("Implied EV at the 75th percentile multiple", at75),
    pad(
      [
        result("Implied EV at the median multiple"),
        blank("result"),
        blank("result"),
        total(impliedEv(STAT_FIRST + STAT_ROW.median), "decimal"),
      ],
      COLUMNS,
      "result",
    ),
  ];
}

export function precedents(): Template {
  return build(
    "precedents",
    "Precedent transactions",
    "Six deals' date, target, acquirer, EV and EBITDA into EV / EBITDA, its six statistics, and the target's implied EV at the 25th percentile, 75th percentile and median multiple.",
    [
      pad([title("Precedent transactions")], COLUMNS, "title"),
      HEADER.map((name) => head(name)),
      ...DEALS.map((deal, index) => dealRow(deal, FIRST_DEAL + index)),
      ...statRows({
        first: FIRST_DEAL,
        last: LAST_DEAL,
        columns: [COL.multiple],
        width: COLUMNS,
      }),
      blankRow(COLUMNS),
      ...impliedRows(),
    ],
  );
}
