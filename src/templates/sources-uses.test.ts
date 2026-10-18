// The sources and uses block against the finance it names: totals that must tie,
// the entry EV off the purchase equity and refinanced net debt, EV / EBITDA and
// the equity share. Computed by test/formula-eval.ts and held to hand figures
// and to a second calculation, on the demo case and on others typed over it.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula, type Template } from "../template-cells";
import { sourcesUses } from "./sources-uses";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// Where things sit with the block's corner at A1: the four sources on rows 3-6
// (amount in B, share in C), their total on row 7, the five uses on rows 10-14,
// their total on row 15, then the check, the entry EV, EBITDA, the multiple and
// the equity share.
const SOURCE_ROWS = [3, 4, 5, 6];
const USE_ROWS = [10, 11, 12, 13, 14];
const CELL = {
  senior: "B3",
  mezzanine: "B4",
  sponsor: "B5",
  rollover: "B6",
  sourcesTotal: "B7",
  sourcesShare: "C7",
  purchase: "B10",
  netDebt: "B11",
  transactionFees: "B12",
  financingFees: "B13",
  cash: "B14",
  usesTotal: "B15",
  usesShare: "C15",
  check: "B17",
  ev: "B19",
  ebitda: "B20",
  multiple: "B21",
  equityShare: "B22",
} as const;

const block = sourcesUses();

// A block cell by its A1 address, the corner being A1.
function cellAt(address: string) {
  const column = address.charCodeAt(0) - 65;
  return block.cells[Number(address.slice(1)) - 1]?.[column];
}

interface Deal {
  senior: number;
  mezzanine: number;
  sponsor: number;
  rollover: number;
  purchase: number;
  netDebt: number;
  transactionFees: number;
  financingFees: number;
  cash: number;
  ebitda: number;
}

const DEMO: Deal = {
  senior: 250,
  mezzanine: 50,
  sponsor: 200,
  rollover: 25,
  purchase: 400,
  netDebt: 100,
  transactionFees: 10,
  financingFees: 5,
  cash: 10,
  ebitda: 62.5,
};

// The same finance written out from the definitions, not from the block's cells.
function reference(deal: Deal) {
  const sources = deal.senior + deal.mezzanine + deal.sponsor + deal.rollover;
  const uses =
    deal.purchase +
    deal.netDebt +
    deal.transactionFees +
    deal.financingFees +
    deal.cash;
  const ev = deal.purchase + deal.netDebt;
  return {
    sources,
    uses,
    ev,
    multiple: deal.ebitda > 0 ? ev / deal.ebitda : "n.m.",
    equityShare: (deal.sponsor + deal.rollover) / sources,
    check: Math.abs(sources - uses) < 0.05 ? "OK" : "CHECK",
  };
}

function overrides(deal: Deal): Record<string, Value> {
  return {
    [CELL.senior]: deal.senior,
    [CELL.mezzanine]: deal.mezzanine,
    [CELL.sponsor]: deal.sponsor,
    [CELL.rollover]: deal.rollover,
    [CELL.purchase]: deal.purchase,
    [CELL.netDebt]: deal.netDebt,
    [CELL.transactionFees]: deal.transactionFees,
    [CELL.financingFees]: deal.financingFees,
    [CELL.cash]: deal.cash,
    [CELL.ebitda]: deal.ebitda,
  };
}

describe("the sources and uses block", () => {
  it("is a 22 by 3 block named Sources and uses, titled by its name", () => {
    expect(block.id).toBe("sources-uses");
    expect(block.name).toBe("Sources and uses");
    expect([block.rows, block.cols]).toEqual([22, 3]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Sources and uses",
    });
  });

  it("labels each line for what it holds, units where there are any", () => {
    expect(block.cells.map((row) => row[0]?.v)).toEqual([
      "Sources and uses",
      "Sources",
      "Senior debt",
      "Mezzanine",
      "Sponsor equity",
      "Management rollover",
      "Total sources",
      undefined,
      "Uses",
      "Purchase equity value",
      "Refinanced net debt",
      "Transaction fees",
      "Financing fees",
      "Cash to balance sheet",
      "Total uses",
      undefined,
      "Check (sources = uses)",
      undefined,
      "Entry EV (EUR m)",
      "LTM EBITDA (EUR m)",
      "Entry EV / EBITDA",
      "Equity % of sources",
    ]);
    expect(block.cells[1]?.map((cell) => cell.v)).toEqual([
      "Sources",
      "EUR m",
      "% of total",
    ]);
    expect(block.cells[8]?.map((cell) => cell.v)).toEqual([
      "Uses",
      "EUR m",
      "% of total",
    ]);
  });

  it("types the nine amounts and the EBITDA in as inputs, EUR millions to one decimal", () => {
    const typed = [
      ...SOURCE_ROWS.map((row) => `B${String(row)}`),
      ...USE_ROWS.map((row) => `B${String(row)}`),
      CELL.ebitda,
    ];
    expect(typed).toHaveLength(10);
    for (const address of typed) {
      expect(cellAt(address), address).toMatchObject({
        kind: "input",
        format: "decimal",
      });
      expect(cellAt(address)?.f, address).toBeUndefined();
    }
  });

  it("keeps sponsor equity a typed input, never the plug that balances the page", () => {
    // A plug would make the two totals tie by construction and the check row
    // prove nothing: the equity cheque is typed, so a page can fail to tie.
    expect(cellAt(CELL.sponsor)?.f).toBeUndefined();
    expect(cellAt(CELL.sponsor)?.v).toBe(200);
  });

  it("adds the sources and the uses to totals that tie in the demo case", () => {
    // by hand: sources 250 + 50 + 200 + 25 = 525;
    //          uses 400 + 100 + 10 + 5 + 10 = 525; the two tie.
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.sourcesTotal)).toBe(525);
    expect(sheet.value(CELL.usesTotal)).toBe(525);
    expect(sheet.value(CELL.check)).toBe("OK");
  });

  it("gives every line its share of its own total, and each side adds to 100 %", () => {
    // by hand: 250 / 525 = 47.62 %, 50 / 525 = 9.52 %, 200 / 525 = 38.10 %,
    //          25 / 525 = 4.76 %; uses 400 / 525 = 76.19 %, 100 / 525 = 19.05 %,
    //          10 / 525 = 1.90 %, 5 / 525 = 0.95 %, 10 / 525 = 1.90 %.
    const sheet = evaluateTemplate(block);
    const sources = [250, 50, 200, 25];
    const uses = [400, 100, 10, 5, 10];
    SOURCE_ROWS.forEach((row, index) => {
      expect(
        sheet.value(`C${String(row)}`),
        `source ${String(row)}`,
      ).toBeCloseTo((sources[index] as number) / 525, 12);
    });
    USE_ROWS.forEach((row, index) => {
      expect(sheet.value(`C${String(row)}`), `use ${String(row)}`).toBeCloseTo(
        (uses[index] as number) / 525,
        12,
      );
    });
    expect(sheet.value(CELL.sourcesShare)).toBeCloseTo(1, 12);
    expect(sheet.value(CELL.usesShare)).toBeCloseTo(1, 12);
  });

  it("builds the entry EV from the purchase equity and the net debt it refinances", () => {
    // by hand: 400 + 100 = 500; fees and cash are uses of funds, not part of EV.
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.ev)).toBe(500);
    const moved = evaluateTemplate(block, {
      [CELL.purchase]: 450,
      [CELL.transactionFees]: 99,
    });
    expect(moved.value(CELL.ev)).toBe(550);
  });

  it("divides the entry EV by the EBITDA into the implied multiple", () => {
    // by hand: 500 / 62.5 = 8.0x.
    expect(evaluateTemplate(block).value(CELL.multiple)).toBe(8);
    expect(
      evaluateTemplate(block, { [CELL.ebitda]: 50 }).value(CELL.multiple),
    ).toBe(10);
  });

  it("says n.m. for the multiple over a loss, a blank or a word, never an error", () => {
    // Text ranks above every number in Excel's comparison, so "n.a." > 0 is TRUE
    // and a bare guard would divide by the word: it reads the cell through N().
    for (const ebitda of [-5, 0, null, "n.a.", "-"]) {
      const sheet = evaluateTemplate(block, { [CELL.ebitda]: ebitda });
      expect(sheet.value(CELL.multiple), String(ebitda)).toBe("n.m.");
    }
  });

  it("takes equity as the sponsor's cheque plus the management rollover over all sources", () => {
    // by hand: (200 + 25) / 525 = 3/7 = 42.857 %.
    expect(evaluateTemplate(block).value(CELL.equityShare)).toBeCloseTo(
      3 / 7,
      12,
    );
  });

  it("reads an absent tranche typed as a word as nothing, not as an error", () => {
    // by hand, mezzanine "n.a." or "-": the sources add 250 + 200 + 25 = 475, the
    // word's own share is 0 and the other three still add to 100 %; the page no
    // longer ties to the 525 of uses, so the check says CHECK. A uses line works
    // the same: financing fees "n.a." leave 520 and a 0 share.
    for (const word of ["n.a.", "-"]) {
      const sheet = evaluateTemplate(block, { [CELL.mezzanine]: word });
      expect(sheet.value(CELL.sourcesTotal), word).toBe(475);
      expect(sheet.value("C4"), word).toBe(0);
      expect(sheet.value(CELL.sourcesShare), word).toBeCloseTo(1, 12);
      expect(sheet.value(CELL.check), word).toBe("CHECK");
      const uses = evaluateTemplate(block, { [CELL.financingFees]: word });
      expect(uses.value(CELL.usesTotal), word).toBe(520);
      expect(uses.value("C13"), word).toBe(0);
      expect(uses.value(CELL.usesShare), word).toBeCloseTo(1, 12);
    }
  });

  it("takes equity over the tranches that exist when the rollover is absent", () => {
    // by hand, rollover "n.a." or "-": sources 250 + 50 + 200 = 500 and equity is
    // the sponsor's 200 alone: 200 / 500 = 40.0 %, not #VALUE!.
    for (const word of ["n.a.", "-"]) {
      const sheet = evaluateTemplate(block, { [CELL.rollover]: word });
      expect(sheet.value(CELL.sourcesTotal), word).toBe(500);
      expect(sheet.value(CELL.equityShare), word).toBeCloseTo(0.4, 12);
    }
  });

  it("builds the EV from the lines that exist when a word stands for none", () => {
    // by hand: net debt "n.a." leaves the purchase equity alone, EV 400.
    const sheet = evaluateTemplate(block, { [CELL.netDebt]: "n.a." });
    expect(sheet.value(CELL.ev)).toBe(400);
  });

  it("reads OK when the totals tie and CHECK when they do not", () => {
    // by hand: sponsor equity 190 leaves sources at 515 against uses of 525,
    // 10 short; 199.96 leaves them 0.04 short, inside the 0.05 tolerance;
    // 199.94 is 0.06 short, outside it.
    const check = (sponsor: number) =>
      evaluateTemplate(block, { [CELL.sponsor]: sponsor }).value(CELL.check);
    expect(check(200)).toBe("OK");
    expect(check(190)).toBe("CHECK");
    expect(check(199.96)).toBe("OK");
    expect(check(199.94)).toBe("CHECK");
    expect(check(210)).toBe("CHECK");
  });

  it.each<[string, Deal]>([
    ["the demo case", DEMO],
    [
      "a heavier equity cheque",
      {
        ...DEMO,
        senior: 200,
        mezzanine: 0,
        sponsor: 300,
        rollover: 25,
      },
    ],
    [
      "a bigger deal with fees at both ends",
      {
        senior: 600,
        mezzanine: 150,
        sponsor: 480,
        rollover: 70,
        purchase: 1050,
        netDebt: 120,
        transactionFees: 22,
        financingFees: 18,
        cash: 90,
        ebitda: 130,
      },
    ],
    ["a page that does not balance", { ...DEMO, sponsor: 150, ebitda: 40 }],
    ["a loss-making target", { ...DEMO, ebitda: -12 }],
  ])("agrees with an independent calculation on %s", (_name, deal) => {
    const sheet = evaluateTemplate(block, overrides(deal));
    const want = reference(deal);
    expect(sheet.value(CELL.sourcesTotal)).toBeCloseTo(want.sources, 9);
    expect(sheet.value(CELL.usesTotal)).toBeCloseTo(want.uses, 9);
    expect(sheet.value(CELL.ev)).toBeCloseTo(want.ev, 9);
    if (typeof want.multiple === "number") {
      expect(sheet.value(CELL.multiple)).toBeCloseTo(want.multiple, 9);
    } else {
      expect(sheet.value(CELL.multiple)).toBe(want.multiple);
    }
    expect(sheet.value(CELL.equityShare)).toBeCloseTo(want.equityShare, 12);
    expect(sheet.value(CELL.check)).toBe(want.check);
  });

  it("shows the divide-by-zero of an empty side instead of a share, and says CHECK", () => {
    const blanks: Record<string, Value> = {};
    for (const row of SOURCE_ROWS) blanks[`B${String(row)}`] = null;
    const sheet = evaluateTemplate(block, blanks);
    expect(sheet.value("C3")).toEqual(new ExcelError("#DIV/0!"));
    expect(sheet.value(CELL.check)).toBe("CHECK");
  });

  it("guards nothing with IFERROR, so no error is ever hidden", () => {
    for (const row of block.cells) {
      for (const cell of row) {
        expect(cell.f ?? "").not.toContain("IFERROR");
      }
    }
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = (block as Template).cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(2, 2)).toBe("=N(D7)/D11");
    expect(formula(9, 2)).toBe("=N(D14)/D19");
    expect(formula(6, 1)).toBe("=SUM(D7:D10)");
    expect(formula(14, 1)).toBe("=SUM(D14:D18)");
    expect(formula(16, 1)).toBe('=IF(ABS(D11-D19)<0.05,"OK","CHECK")');
    expect(formula(18, 1)).toBe("=SUM(D14:D15)");
    expect(formula(20, 1)).toBe('=IF(N(D24)>0,D23/D24,"n.m.")');
    expect(formula(21, 1)).toBe("=SUM(D9:D10)/D11");
  });
});
