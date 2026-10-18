// The returns block against the finance it names: exit equity (EBITDA times a
// multiple less net debt), MOIC, IRR over a cash-flow row with its sign change
// built in, and MOIC a multiple step either side. Computed by test/formula-eval.ts
// and held to hand figures, the closed-form IRR and the NPV that IRR must zero.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { returns } from "./returns";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: the five assumptions and the exit equity sit in
// B2:B7, the year header on row 9, the equity cash flow on row 10 (year 0 in B,
// year 5 in G), MOIC and IRR in B12 and B13, the step in B15, the three exit
// multiples in B16:D16 and the MOIC at each in B17:D17.
const CELL = {
  entry: "B2",
  ebitda: "B3",
  multiple: "B4",
  ev: "B5",
  netDebt: "B6",
  exitEquity: "B7",
  moic: "B12",
  irr: "B13",
  step: "B15",
} as const;
const FLOW_ROW = 10;
const YEAR_COLUMNS = ["B", "C", "D", "E", "F", "G"];
const flowCell = (year: number): string =>
  `${YEAR_COLUMNS[year] as string}${String(FLOW_ROW)}`;
const SENSITIVITY = {
  multiples: ["B16", "C16", "D16"],
  moic: ["B17", "C17", "D17"],
};
const YEARS = 5;

const block = returns();

interface Case {
  entry: number;
  ebitda: number;
  multiple: number;
  netDebt: number;
}

const DEMO: Case = { entry: 225, ebitda: 92, multiple: 8, netDebt: 175 };

function overrides(deal: Case): Record<string, Value> {
  return {
    [CELL.entry]: deal.entry,
    [CELL.ebitda]: deal.ebitda,
    [CELL.multiple]: deal.multiple,
    [CELL.netDebt]: deal.netDebt,
  };
}

// The same finance from the definitions: exit equity is EBITDA times the multiple
// less net debt, MOIC is it over the equity in, and with nothing in between the
// IRR has a closed form, the fifth root of the MOIC less one.
function reference(deal: Case) {
  const exitEquity = deal.ebitda * deal.multiple - deal.netDebt;
  const moic = exitEquity / deal.entry;
  const irr = exitEquity > 0 ? moic ** (1 / YEARS) - 1 : "n/a";
  return { exitEquity, moic, irr };
}

// The net present value at a rate, written out here so the IRR the sheet shows is
// checked against the definition of an IRR and not against another solver.
const npv = (rate: number, flows: number[]): number =>
  flows.reduce((sum, flow, year) => sum + flow / (1 + rate) ** year, 0);

describe("the returns block", () => {
  it("is a 17 by 7 block named Returns (MOIC / IRR), titled by its name", () => {
    expect(block.id).toBe("returns");
    expect(block.name).toBe("Returns (MOIC / IRR)");
    expect([block.rows, block.cols]).toEqual([17, 7]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Returns (MOIC / IRR)",
    });
  });

  it("labels each row for what it holds, and names the years 0 to 5 in the header", () => {
    expect(block.cells.map((row) => row[0]?.v)).toEqual([
      "Returns (MOIC / IRR)",
      "Entry equity (EUR m)",
      "Exit EBITDA (EUR m)",
      "Exit multiple",
      "Exit EV (EUR m)",
      "Net debt at exit (EUR m)",
      "Exit equity (EUR m)",
      undefined,
      "EUR m",
      "Equity cash flow (EUR m)",
      undefined,
      "MOIC",
      "IRR",
      undefined,
      "Multiple step",
      "Exit multiple",
      "MOIC at that multiple",
    ]);
    const header = block.cells[8];
    expect(header?.every((found) => found.kind === "header")).toBe(true);
    // Words, not numbers: the number formats show a typed 0 as a dash, and year 0
    // is the one header that must read "0".
    expect(header?.slice(1).map((found) => found.v)).toEqual([
      "Year 0",
      "Year 1",
      "Year 2",
      "Year 3",
      "Year 4",
      "Year 5",
    ]);
  });

  it("types the five assumptions in as inputs, EUR millions to one decimal and the multiples as multiples", () => {
    const at = (address: string) =>
      block.cells[Number(address.slice(1)) - 1]?.[
        "ABCDEFG".indexOf(address[0] as string)
      ];
    const typed: [string, string][] = [
      [CELL.entry, "decimal"],
      [CELL.ebitda, "decimal"],
      [CELL.multiple, "multiple"],
      [CELL.netDebt, "decimal"],
      [CELL.step, "multiple"],
    ];
    for (const [address, format] of typed) {
      expect(at(address), address).toMatchObject({ kind: "input", format });
      expect(at(address)?.f, address).toBeUndefined();
    }
    for (const address of [CELL.ev, CELL.exitEquity]) {
      expect(at(address), address).toMatchObject({
        kind: "formula",
        format: "decimal",
      });
    }
    expect(at(CELL.moic)).toMatchObject({ kind: "result", format: "multiple" });
    expect(at(CELL.irr)).toMatchObject({ kind: "result", format: "percent" });
  });

  it("builds the exit equity from EBITDA times the multiple less net debt", () => {
    // by hand: EV 92.0 x 8.0 = 736.0; less net debt 175.0 = exit equity 561.0.
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.ev)).toBe(736);
    expect(sheet.value(CELL.exitEquity)).toBe(561);
  });

  it("lays the cash flows out with the sign change built in: entry out, zeros, exit in", () => {
    // by hand: year 0 is minus the 225.0 put in, years 1 to 4 are zero, year 5 is
    // the 561.0 exit equity.
    const sheet = evaluateTemplate(block);
    expect(sheet.value(flowCell(0))).toBe(-225);
    for (let year = 1; year < YEARS; year += 1) {
      expect(sheet.value(flowCell(year)), `year ${String(year)}`).toBe(0);
    }
    expect(sheet.value(flowCell(YEARS))).toBe(561);
    const flows = Array.from(
      { length: YEARS + 1 },
      (_unused, year) => sheet.value(flowCell(year)) as number,
    );
    expect(Math.min(...flows)).toBeLessThan(0);
    expect(Math.max(...flows)).toBeGreaterThan(0);
  });

  it("types the four middle years as zeros, because Excel's IRR skips a blank", () => {
    for (let year = 1; year < YEARS; year += 1) {
      const found = block.cells[FLOW_ROW - 1]?.[year + 1];
      expect(found, `year ${String(year)}`).toMatchObject({
        kind: "input",
        v: 0,
        format: "decimal",
      });
    }
  });

  it("answers n/a, never a rate on a squeezed timeline, when a year is cleared or not a number", () => {
    // Excel's IRR skips a blank or a word, so five years read as four or two and
    // the rate looks fine. by hand: year 2 cleared reads -225, 0, 0, 0, 561 one
    // period short, 25.66 % instead of 20.05 %; all four zeros cleared read
    // 561 / 225 - 1 = 149.3 %. The block says n/a for each instead.
    const cleared = (years: number[]): Record<string, Value> =>
      Object.fromEntries(years.map((year) => [flowCell(year), null]));
    expect(evaluateTemplate(block, cleared([2])).value(CELL.irr)).toBe("n/a");
    expect(evaluateTemplate(block, cleared([1, 2, 3, 4])).value(CELL.irr)).toBe(
      "n/a",
    );
    const word = evaluateTemplate(block, { [flowCell(2)]: "n.a." });
    expect(word.value(CELL.irr)).toBe("n/a");
    // The MOIC is untouched by any of it: a blank or a word adds nothing.
    expect(word.value(CELL.moic)).toBeCloseTo(561 / 225, 12);
  });

  it("adds every inflow and divides by the entry equity for MOIC", () => {
    // by hand: nothing in between, so 561.0 / 225.0 = 2.4933x.
    expect(evaluateTemplate(block).value(CELL.moic)).toBeCloseTo(561 / 225, 12);
    expect(evaluateTemplate(block).value(CELL.moic)).toBeCloseTo(2.4933, 4);
  });

  it("finds the IRR in the cash-flow row, 20.05 % on the demo", () => {
    // by hand: 561 / 225 = 2.4933x over five years with nothing in between, so
    // (1 + r)^5 = 2.4933 and r = 2.4933^(1/5) - 1 = 0.20048 (20.0 %); and the
    // flows -225, 0, 0, 0, 0, 561 discounted at that rate come to zero.
    const irr = evaluateTemplate(block).value(CELL.irr) as number;
    expect(irr).toBeCloseTo(0.20048314988, 10);
    expect(irr).toBeCloseTo((561 / 225) ** 0.2 - 1, 12);
    expect(npv(irr, [-225, 0, 0, 0, 0, 561])).toBeCloseTo(0, 9);
  });

  it("shows the MOIC at the exit multiple one step below and above, the middle one the MOIC itself", () => {
    // by hand: at 7.0x EV is 644.0, equity 469.0, MOIC 469 / 225 = 2.0844x; at
    // 8.0x 561 / 225 = 2.4933x; at 9.0x EV 828.0, equity 653.0, MOIC 2.9022x.
    const sheet = evaluateTemplate(block);
    expect(
      SENSITIVITY.multiples.map((address) => sheet.value(address)),
    ).toEqual([7, 8, 9]);
    const [low, base, high] = SENSITIVITY.moic.map(
      (address) => sheet.value(address) as number,
    );
    expect(low).toBeCloseTo(469 / 225, 12);
    expect(base).toBeCloseTo(561 / 225, 12);
    expect(high).toBeCloseTo(653 / 225, 12);
    expect(base).toBe(sheet.value(CELL.moic));
  });

  it("moves the sensitivity with the step and the base multiple typed over", () => {
    // by hand, step 2.0x on a 10.0x base: 8.0x, 10.0x, 12.0x; equity 92 x 8 - 175 =
    // 561 / 225; 92 x 10 - 175 = 745; 92 x 12 - 175 = 929.
    const sheet = evaluateTemplate(block, {
      [CELL.multiple]: 10,
      [CELL.step]: 2,
    });
    expect(
      SENSITIVITY.multiples.map((address) => sheet.value(address)),
    ).toEqual([8, 10, 12]);
    const moic = SENSITIVITY.moic.map(
      (address) => sheet.value(address) as number,
    );
    expect(moic[0]).toBeCloseTo(561 / 225, 12);
    expect(moic[1]).toBeCloseTo(745 / 225, 12);
    expect(moic[2]).toBeCloseTo(929 / 225, 12);
  });

  it.each<[string, Case]>([
    ["the demo case", DEMO],
    [
      "a bigger win on a smaller cheque",
      { entry: 100, ebitda: 50, multiple: 10, netDebt: 200 },
    ],
    [
      "a loss that still returns some equity",
      { entry: 400, ebitda: 60, multiple: 7, netDebt: 380 },
    ],
    [
      "no multiple expansion on flat debt",
      { entry: 300, ebitda: 70, multiple: 9, netDebt: 330 },
    ],
  ])("agrees with the definitions on %s", (_name, deal) => {
    // by hand, "a loss": 60 x 7.0 = 420, less 380 = 40.0 back on 400.0, MOIC
    // 0.10x and IRR 0.1^(1/5) - 1 = -36.9 %: negative, and still an IRR.
    const sheet = evaluateTemplate(block, overrides(deal));
    const want = reference(deal);
    expect(sheet.value(CELL.exitEquity)).toBeCloseTo(want.exitEquity, 9);
    expect(sheet.value(CELL.moic)).toBeCloseTo(want.moic, 12);
    const irr = sheet.value(CELL.irr) as number;
    expect(irr).toBeCloseTo(want.irr as number, 10);
    expect(npv(irr, [-deal.entry, 0, 0, 0, 0, want.exitEquity])).toBeCloseTo(
      0,
      6,
    );
  });

  it("answers n/a, not an error, when the exit leaves no equity", () => {
    // by hand: 92 x 8.0 = 736.0 less net debt 800.0 = -64.0: the row is -225.0
    // then zeros then -64.0, one-signed, so IRR has nothing to solve; the MOIC
    // still shows what it is, -64 / 225 = -0.28x.
    const sheet = evaluateTemplate(block, { [CELL.netDebt]: 800 });
    expect(sheet.value(CELL.exitEquity)).toBe(-64);
    expect(sheet.value(CELL.irr)).toBe("n/a");
    expect(sheet.value(CELL.moic)).toBeCloseTo(-64 / 225, 12);
    const wiped = evaluateTemplate(block, { [CELL.netDebt]: 736 });
    expect(wiped.value(CELL.irr)).toBe("n/a");
  });

  it("answers n/a when the entry equity is typed negative, and #DIV/0! for a MOIC on none", () => {
    // A negative entry puts a positive number in year 0 and leaves no outflow.
    expect(
      evaluateTemplate(block, { [CELL.entry]: -225 }).value(CELL.irr),
    ).toBe("n/a");
    const none = evaluateTemplate(block, { [CELL.entry]: 0 });
    expect(none.value(CELL.irr)).toBe("n/a");
    expect(none.value(CELL.moic)).toEqual(new ExcelError("#DIV/0!"));
  });

  it("guards exactly one cell with IFERROR: the IRR, and nothing else", () => {
    const guarded: string[] = [];
    block.cells.forEach((row, r) =>
      row.forEach((found, c) => {
        if (found.f?.includes("IFERROR"))
          guarded.push(`r${String(r)}c${String(c)}`);
      }),
    );
    expect(guarded).toEqual(["r12c1"]);
    const formula = block.cells[12]?.[1]?.f ?? "";
    expect(formula.match(/IFERROR/g)).toHaveLength(1);
    // The squeezed-timeline guard, then IRR with a guess at its own root: the
    // MOIC's fifth root less one, or Excel's 10 % when the MOIC is not positive.
    expect(resolveFormula(formula, { row: 0, column: 0 })).toBe(
      '=IFERROR(IF(COUNT(B10:G10)<6,"n/a",IRR(B10:G10,IF(N(B12)>0,B12^(1/5)-1,0.1))),"n/a")',
    );
  });

  it("lets an error in any other cell show: a text multiple is #VALUE! all the way through", () => {
    const sheet = evaluateTemplate(block, { [CELL.multiple]: "8x" });
    expect(sheet.value(CELL.ev)).toEqual(new ExcelError("#VALUE!"));
    expect(sheet.value(CELL.moic)).toEqual(new ExcelError("#VALUE!"));
    expect(sheet.value(CELL.irr)).toBe("n/a");
  });

  it("falls back to the default guess when the MOIC is not positive but the cash still has a rate", () => {
    // by hand: 440.0 back in year 1 and an exit worth -450.0 (net debt 1 186.0
    // against EV 736.0) on 225.0 in is a MOIC of (440 - 450) / 225 = -0.04x, so
    // MOIC^(1/5) has no real value; the guess falls back to 10 % and the row
    // -225, 440, 0, 0, 0, -450 still has a root, 34.57 %.
    const sheet = evaluateTemplate(block, {
      [CELL.netDebt]: 1186,
      [flowCell(1)]: 440,
    });
    expect(sheet.value(CELL.exitEquity)).toBe(-450);
    expect(sheet.value(CELL.moic) as number).toBeLessThan(0);
    const irr = sheet.value(CELL.irr) as number;
    expect(irr).toBeCloseTo(0.3457172809, 8);
    expect(npv(irr, [-225, 440, 0, 0, 0, -450])).toBeCloseTo(0, 8);
  });

  it("counts a dividend typed into the middle years in the MOIC and the IRR alike", () => {
    // by hand: MOIC is every euro back over the equity in, 30.0 in year 3 plus the
    // 561.0 exit is 591.0 on 225.0 = 2.6267x; IRR reads the same row, so 30.0 in
    // year 3 lifts it from 20.05 % to 21.90 % (the root of
    // -225 + 30 / (1 + r)^3 + 561 / (1 + r)^5 = 0).
    const sheet = evaluateTemplate(block, { [flowCell(3)]: 30 });
    expect(sheet.value(CELL.moic)).toBeCloseTo(591 / 225, 12);
    expect(sheet.value(CELL.moic)).toBeCloseTo(2.6267, 4);
    const irr = sheet.value(CELL.irr) as number;
    expect(irr).toBeCloseTo(0.2189824389, 8);
    expect(npv(irr, [-225, 0, 0, 30, 0, 561])).toBeCloseTo(0, 8);
  });

  it("carries the dividend into the sensitivity, so its middle column stays the MOIC", () => {
    // by hand, 30.0 in year 3: at 7.0x (30 + 644 - 175) / 225 = 499 / 225 = 2.2178x;
    // at 8.0x (30 + 736 - 175) / 225 = 591 / 225 = 2.6267x; at 9.0x
    // (30 + 828 - 175) / 225 = 683 / 225 = 3.0356x.
    const sheet = evaluateTemplate(block, { [flowCell(3)]: 30 });
    const [low, base, high] = SENSITIVITY.moic.map(
      (address) => sheet.value(address) as number,
    );
    expect(low).toBeCloseTo(499 / 225, 12);
    expect(base).toBeCloseTo(591 / 225, 12);
    expect(high).toBeCloseTo(683 / 225, 12);
    expect(base).toBeCloseTo(sheet.value(CELL.moic) as number, 12);
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const found = block.cells[row]?.[column];
      return found?.f === undefined ? undefined : resolveFormula(found.f, C5);
    };
    expect(formula(4, 1)).toBe("=D7*D8");
    expect(formula(6, 1)).toBe("=D9-D10");
    expect(formula(9, 1)).toBe("=-D6");
    expect(formula(9, 6)).toBe("=D11");
    expect(formula(11, 1)).toBe("=SUM(E14:I14)/D6");
    expect(formula(12, 1)).toBe(
      '=IFERROR(IF(COUNT(D14:I14)<6,"n/a",IRR(D14:I14,IF(N(D16)>0,D16^(1/5)-1,0.1))),"n/a")',
    );
    expect(formula(15, 1)).toBe("=D8-D19");
    expect(formula(15, 2)).toBe("=D8");
    expect(formula(15, 3)).toBe("=D8+D19");
    expect(formula(16, 1)).toBe("=(SUM(E14:H14)+D7*D20-D10)/D6");
    expect(formula(16, 3)).toBe("=(SUM(E14:H14)+D7*F20-D10)/D6");
  });
});
