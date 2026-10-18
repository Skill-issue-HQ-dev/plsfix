// The sources and uses block: four sources and five uses of funds in EUR
// millions with their shares of the total, a check that the two totals tie, the
// entry EV and its EV / EBITDA, and equity as a share of the sources. Invariant:
// sponsor equity is typed, never the plug, so a page that is out says CHECK.

import {
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
import { at, span } from "./refs";

const COLUMNS = 3;
// Amounts sit in column 1 and each line's share of its total in column 2.
const AMOUNT = 1;
const SHARE = 2;
// The check's tolerance in EUR millions: a page 50 000 out does not tie.
const TIE_TOLERANCE = 0.05;

type Line = [name: string, amount: number];

// EUR millions: a mid-market buyout at 8.0x EBITDA, 4.8x levered, 43 % equity.
const SOURCES: Line[] = [
  ["Senior debt", 250],
  ["Mezzanine", 50],
  ["Sponsor equity", 200],
  ["Management rollover", 25],
];
const USES: Line[] = [
  ["Purchase equity value", 400],
  ["Refinanced net debt", 100],
  ["Transaction fees", 10],
  ["Financing fees", 5],
  ["Cash to balance sheet", 10],
];
const LTM_EBITDA = 62.5;

// Rows, top to bottom: the title, the sources and their total, a gap, the uses
// and their total, a gap, the check, a gap, then the entry EV and its ratios.
const SOURCES_HEADER = 1;
const SOURCES_FIRST = SOURCES_HEADER + 1;
const SOURCES_LAST = SOURCES_FIRST + SOURCES.length - 1;
const SOURCES_TOTAL = SOURCES_LAST + 1;
const USES_HEADER = SOURCES_TOTAL + 2;
const USES_FIRST = USES_HEADER + 1;
const USES_LAST = USES_FIRST + USES.length - 1;
const USES_TOTAL = USES_LAST + 1;
const CHECK = USES_TOTAL + 2;
const ENTRY_EV = CHECK + 2;
const EBITDA = ENTRY_EV + 1;
// Equity is the sponsor's cheque and the management rollover: the third and
// fourth sources. The EV is the equity bought plus the net debt refinanced: the
// first two uses (fees and cash are uses of funds, not part of the price). Each
// pair is summed over its two adjacent lines, so a tranche typed as a word to
// say it is absent ("n.a.", "-") counts as nothing instead of breaking the sum.
const SPONSOR_ROW = SOURCES_FIRST + 2;
const ROLLOVER_ROW = SOURCES_FIRST + 3;
const PURCHASE_ROW = USES_FIRST;
const REFINANCED_ROW = USES_FIRST + 1;

const side = (name: string): TemplateCell[] =>
  [name, "EUR m", "% of total"].map((text) => head(text));

// A line's share reads the amount through N(), so a word typed for an absent
// tranche is a 0 share, not a #VALUE! that runs into the column's total.
function lineRow(
  [name, amount]: Line,
  row: number,
  totalRow: number,
): TemplateCell[] {
  return [
    label(name),
    input(amount, "decimal"),
    calc(`=N(${at(row, AMOUNT)})/${at(totalRow, AMOUNT)}`, "percent"),
  ];
}

function totalRow(name: string, first: number, last: number): TemplateCell[] {
  const sum = (column: number): string =>
    `=SUM(${span(first, column, last, column)})`;
  return [
    result(name),
    total(sum(AMOUNT), "decimal"),
    total(sum(SHARE), "percent"),
  ];
}

// The page ties when the two totals agree to within the tolerance.
function checkRow(): TemplateCell[] {
  const gap = `ABS(${at(SOURCES_TOTAL, AMOUNT)}-${at(USES_TOTAL, AMOUNT)})`;
  return pad(
    [
      label("Check (sources = uses)"),
      calc(`=IF(${gap}<${String(TIE_TOLERANCE)},"OK","CHECK")`, "decimal"),
    ],
    COLUMNS,
  );
}

// A multiple over a zero, negative or typed-word EBITDA is not meaningful: it says
// n.m. N() makes a word a 0, because text ranks above every number in a comparison
// and "n.a." > 0 would let the guard divide by the word.
function entryRows(): TemplateCell[][] {
  const ev = at(ENTRY_EV, AMOUNT);
  const ebitda = at(EBITDA, AMOUNT);
  const equity = span(SPONSOR_ROW, AMOUNT, ROLLOVER_ROW, AMOUNT);
  const bought = span(PURCHASE_ROW, AMOUNT, REFINANCED_ROW, AMOUNT);
  return [
    [label("Entry EV (EUR m)"), calc(`=SUM(${bought})`, "decimal")],
    [label("LTM EBITDA (EUR m)"), input(LTM_EBITDA, "decimal")],
    [
      label("Entry EV / EBITDA"),
      calc(`=IF(N(${ebitda})>0,${ev}/${ebitda},"n.m.")`, "multiple"),
    ],
    [
      label("Equity % of sources"),
      calc(`=SUM(${equity})/${at(SOURCES_TOTAL, AMOUNT)}`, "percent"),
    ],
  ].map((cells) => pad(cells, COLUMNS));
}

export function sourcesUses(): Template {
  return build(
    "sources-uses",
    "Sources and uses",
    "Four sources and five uses of funds in EUR millions with their shares of the total, a check that the two totals tie, the entry EV and its EV / EBITDA, and equity (sponsor plus rollover) as a share of the sources.",
    [
      pad([title("Sources and uses")], COLUMNS, "title"),
      side("Sources"),
      ...SOURCES.map((line, index) =>
        lineRow(line, SOURCES_FIRST + index, SOURCES_TOTAL),
      ),
      totalRow("Total sources", SOURCES_FIRST, SOURCES_LAST),
      blankRow(COLUMNS),
      side("Uses"),
      ...USES.map((line, index) =>
        lineRow(line, USES_FIRST + index, USES_TOTAL),
      ),
      totalRow("Total uses", USES_FIRST, USES_LAST),
      blankRow(COLUMNS),
      checkRow(),
      blankRow(COLUMNS),
      ...entryRows(),
    ],
  );
}
