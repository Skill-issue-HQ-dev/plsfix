// The debt paydown block against the finance it names: interest on the opening
// balance, amortisation of the original loan, a cash sweep and leverage. Computed
// by test/formula-eval.ts and held to a schedule worked out by hand, to a second
// implementation, to edge cases and to 200 seeded plans.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { debtPaydown } from "./debt-paydown";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1 the four inputs sit in B2:B5, the year header on
// row 7, CFADS and EBITDA on rows 8 and 9 and the schedule on rows 11 to 16, the
// five years in columns B to F; leverage is row 18.
const INPUT = { loan: "B2", rate: "B3", amortisation: "B4", sweep: "B5" };
const ROW = {
  cfads: 8,
  ebitda: 9,
  opening: 11,
  interest: 12,
  mandatory: 13,
  available: 14,
  sweep: 15,
  closing: 16,
  leverage: 18,
} as const;
const YEARS = 5;
const cell = (row: number, year: number): string =>
  `${"BCDEF"[year - 1] as string}${String(row)}`;

type Sheet = ReturnType<typeof evaluateTemplate>;
const num = (sheet: Sheet, row: number, year: number): number =>
  sheet.value(cell(row, year)) as number;
const years = Array.from({ length: YEARS }, (_unused, index) => index + 1);

interface Plan {
  loan: number;
  rate: number;
  amortisation: number;
  sweep: number;
  cfads: number[];
  ebitda: number[];
}

const DEMO: Plan = {
  loan: 250,
  rate: 0.07,
  amortisation: 0.05,
  sweep: 0.5,
  cfads: [40, 44, 48, 52, 56],
  ebitda: [68, 74, 80, 86, 92],
};

const FIELDS = [
  "opening",
  "interest",
  "mandatory",
  "available",
  "sweep",
  "closing",
] as const;
type Year = Record<(typeof FIELDS)[number], number>;

// The schedule from the definitions, as a loop: interest on what is owed at the
// start of the year, the mandatory slice of the ORIGINAL loan (never more than
// is owed), the sweep share of whatever cash is left after interest and that
// slice (never negative, never more than the balance still owed).
function reference(plan: Plan): Year[] {
  let opening = plan.loan;
  return plan.cfads.map((cfads) => {
    const interest = opening * plan.rate;
    const mandatory = Math.min(plan.amortisation * plan.loan, opening);
    const available = cfads - interest - mandatory;
    const sweep = Math.min(
      plan.sweep * Math.max(0, available),
      opening - mandatory,
    );
    const closing = opening - mandatory - sweep;
    const year = { opening, interest, mandatory, available, sweep, closing };
    opening = closing;
    return year;
  });
}

function overrides(plan: Plan): Record<string, Value> {
  const values: Record<string, Value> = {
    [INPUT.loan]: plan.loan,
    [INPUT.rate]: plan.rate,
    [INPUT.amortisation]: plan.amortisation,
    [INPUT.sweep]: plan.sweep,
  };
  for (const year of years) {
    values[cell(ROW.cfads, year)] = plan.cfads[year - 1] as number;
    values[cell(ROW.ebitda, year)] = plan.ebitda[year - 1] as number;
  }
  return values;
}

// Holds every schedule cell of a computed sheet to the second implementation.
function expectSchedule(sheet: Sheet, plan: Plan, where = ""): void {
  reference(plan).forEach((want, index) => {
    for (const field of FIELDS) {
      const named = `${where}${field}, year ${String(index + 1)}`;
      expect(num(sheet, ROW[field], index + 1), named).toBeCloseTo(
        want[field],
        9,
      );
    }
  });
}

// A fixed-seed linear congruential generator, so a failure reproduces.
function seeded(seed: number): () => number {
  let state = seed;
  return () => {
    state = (state * 1_664_525 + 1_013_904_223) % 4_294_967_296;
    return state / 4_294_967_296;
  };
}

const block = debtPaydown();

describe("the debt paydown block", () => {
  it("is an 18 by 6 block named Debt paydown (cash sweep), titled by its name", () => {
    expect(block.id).toBe("debt-paydown");
    expect(block.name).toBe("Debt paydown (cash sweep)");
    expect([block.rows, block.cols]).toEqual([18, 6]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Debt paydown (cash sweep)",
    });
  });

  it("labels each row for what it holds and numbers the five years in the header", () => {
    expect(block.cells.map((row) => row[0]?.v)).toEqual([
      "Debt paydown (cash sweep)",
      "Term loan at close (EUR m)",
      "Interest rate",
      "Mandatory amortisation (% of original)",
      "Cash sweep (% of excess cash)",
      undefined,
      "Year (EUR m)",
      "CFADS",
      "EBITDA",
      undefined,
      "Opening balance",
      "Interest",
      "Mandatory amortisation",
      "Cash available for sweep",
      "Cash sweep",
      "Closing balance",
      undefined,
      "Leverage (closing debt / EBITDA)",
    ]);
    const header = block.cells[6];
    expect(header?.every((found) => found.kind === "header")).toBe(true);
    expect(header?.slice(1).map((found) => found.v)).toEqual(years);
  });

  it("types the four assumptions and the ten yearly figures in as inputs, in the formats a banker reads them", () => {
    const at = (address: string) =>
      block.cells[Number(address.slice(1)) - 1]?.[
        "ABCDEF".indexOf(address[0] as string)
      ];
    const typed: [string, string][] = [
      [INPUT.loan, "decimal"],
      [INPUT.rate, "percent"],
      [INPUT.amortisation, "percent"],
      [INPUT.sweep, "percent"],
    ];
    for (const year of years) {
      for (const row of [ROW.cfads, ROW.ebitda]) {
        typed.push([cell(row, year), "decimal"]);
      }
    }
    expect(typed).toHaveLength(14);
    for (const [address, format] of typed) {
      expect(at(address), address).toMatchObject({ kind: "input", format });
      expect(at(address)?.f, address).toBeUndefined();
    }
  });

  it("makes every schedule row a formula, the closing balance the result row, leverage a multiple", () => {
    const rows = [ROW.opening, ROW.interest, ROW.mandatory, ROW.available];
    for (const year of years) {
      for (const row of [...rows, ROW.sweep]) {
        const found = block.cells[row - 1]?.[year];
        expect(found, cell(row, year)).toMatchObject({
          kind: "formula",
          format: "decimal",
        });
        expect(found?.v, cell(row, year)).toBeUndefined();
      }
      expect(block.cells[ROW.closing - 1]?.[year]).toMatchObject({
        kind: "result",
        format: "decimal",
      });
      expect(block.cells[ROW.leverage - 1]?.[year]).toMatchObject({
        kind: "formula",
        format: "multiple",
      });
    }
  });

  it("works the demo loan down year by year", () => {
    // by hand, the demo (EUR m): loan 250.0 at 7.0 %, 12.5 (5 % of it) amortised a
    // year, half the cash left swept; CFADS 40, 44, 48, 52, 56. Year 1: interest
    // 250 x 7 % = 17.5, cash left 40 - 17.5 - 12.5 = 10.0, sweep 5.0, closing 232.5;
    // year 2: interest 16.275, cash left 15.225, sweep 7.6125, closing 212.3875; the
    // rest the same way, down to 135.4445114265625 after year 5.
    const byHand = {
      interest: [17.5, 16.275, 14.867125, 13.269974375, 11.476923478125],
      available: [10, 15.225, 20.632875, 26.230025625, 32.023076521875],
      sweep: [5, 7.6125, 10.3164375, 13.1150128125, 16.0115382609375],
      closing: [
        232.5, 212.3875, 189.5710625, 163.9560496875, 135.4445114265625,
      ],
    };
    const sheet = evaluateTemplate(block);
    for (const [field, values] of Object.entries(byHand)) {
      values.forEach((want, index) => {
        const where = `${field}, year ${String(index + 1)}`;
        const row = ROW[field as keyof typeof byHand];
        expect(num(sheet, row, index + 1), where).toBeCloseTo(want, 9);
      });
    }
  });

  it("divides each closing balance by that year's EBITDA for leverage", () => {
    // by hand: 232.5 / 68 = 3.4191x in year 1 and 135.4445 / 92 = 1.4722x in year 5.
    const sheet = evaluateTemplate(block);
    const year5 = 135.4445114265625 / 92;
    expect(num(sheet, ROW.leverage, 1)).toBeCloseTo(232.5 / 68, 12);
    expect(num(sheet, ROW.leverage, 1)).toBeCloseTo(3.4191, 4);
    expect(num(sheet, ROW.leverage, 5)).toBeCloseTo(year5, 12);
    expect(num(sheet, ROW.leverage, 5)).toBeCloseTo(1.4722, 4);
  });

  it("says n.m. for leverage over a zero, negative, blank or typed-word EBITDA, never an error", () => {
    // Text ranks above every number in Excel's comparison, so "n.a." > 0 is TRUE and
    // a bare guard would divide by the word: the guard reads EBITDA through N().
    for (const ebitda of [0, -10, null, "n.a.", "-"]) {
      const sheet = evaluateTemplate(block, { [cell(ROW.ebitda, 3)]: ebitda });
      expect(sheet.value(cell(ROW.leverage, 3))).toBe("n.m.");
      expect(sheet.value(cell(ROW.leverage, 2))).not.toBe("n.m.");
    }
  });

  it("takes interest off the opening balance, so this year's cash never changes this year's interest", () => {
    // Interest on the opening balance is what keeps the block free of any
    // circular reference: more year-1 cash sweeps more, moving year 1's CLOSING
    // balance and every interest charge after it, but not year 1's own interest.
    const base = evaluateTemplate(block);
    const richer = evaluateTemplate(block, { [cell(ROW.cfads, 1)]: 80 });
    expect(num(richer, ROW.interest, 1)).toBe(num(base, ROW.interest, 1));
    expect(num(richer, ROW.opening, 1)).toBe(250);
    expect(num(richer, ROW.closing, 1)).toBeLessThan(num(base, ROW.closing, 1));
    expect(num(richer, ROW.interest, 2)).toBeLessThan(
      num(base, ROW.interest, 2),
    );
  });

  it("amortises a fixed slice of the ORIGINAL loan, not of the shrinking balance", () => {
    // by hand: 5 % of 250 = 12.5 in all five years; 5 % of year 5's opening
    // balance (163.96) would be 8.2, which is not what a term loan's schedule says.
    const sheet = evaluateTemplate(block);
    for (const year of years) {
      expect(
        num(sheet, ROW.mandatory, year),
        `year ${String(year)}`,
      ).toBeCloseTo(12.5, 12);
    }
  });

  it("sweeps every euro of what is left at 100 % and nothing at 0 %", () => {
    // by hand, 100 %: year 1 sweeps all of the 10.0 left, closing 250 - 12.5 - 10 = 227.5.
    const full = evaluateTemplate(block, { [INPUT.sweep]: 1 });
    expect(num(full, ROW.sweep, 1)).toBeCloseTo(10, 12);
    expect(num(full, ROW.closing, 1)).toBeCloseTo(227.5, 12);
    // by hand, 0 %: only the 12.5 a year amortises, 250 - 5 x 12.5 = 187.5.
    const none = evaluateTemplate(block, { [INPUT.sweep]: 0 });
    for (const year of years) expect(num(none, ROW.sweep, year)).toBe(0);
    expect(num(none, ROW.closing, 5)).toBeCloseTo(187.5, 12);
  });

  it("sweeps nothing and lets the shortfall show when the cash does not cover interest and amortisation", () => {
    // by hand, CFADS 20 against interest 17.5 and amortisation 12.5 in year 1:
    // 20 - 17.5 - 12.5 = -10.0 left, shown as it is; sweep max(0, -10) = 0;
    // closing 250 - 12.5 = 237.5. The balance never grows to pay a shortfall.
    const sheet = evaluateTemplate(block, { [cell(ROW.cfads, 1)]: 20 });
    expect(num(sheet, ROW.available, 1)).toBeCloseTo(-10, 12);
    expect(num(sheet, ROW.sweep, 1)).toBe(0);
    expect(num(sheet, ROW.closing, 1)).toBeCloseTo(237.5, 12);
  });

  it("stops at zero: a small loan is swept out in year 1 and nothing is charged after", () => {
    // by hand: loan 30.0, 100 % sweep: interest 2.1, mandatory 5 % x 30 = 1.5,
    // cash left 40 - 2.1 - 1.5 = 36.4, sweep min(36.4, 30 - 1.5 = 28.5) = 28.5,
    // closing 30 - 1.5 - 28.5 = 0; from year 2 on there is nothing to charge or repay,
    // and the mandatory slice min(1.5, 0) = 0 is capped at what is owed.
    const sheet = evaluateTemplate(block, {
      [INPUT.loan]: 30,
      [INPUT.sweep]: 1,
    });
    expect(num(sheet, ROW.interest, 1)).toBeCloseTo(2.1, 12);
    expect(num(sheet, ROW.mandatory, 1)).toBeCloseTo(1.5, 12);
    expect(num(sheet, ROW.sweep, 1)).toBeCloseTo(28.5, 12);
    for (const year of years) {
      expect(num(sheet, ROW.closing, year)).toBeCloseTo(0, 12);
    }
    for (const year of years.slice(1)) {
      for (const row of [ROW.interest, ROW.mandatory, ROW.sweep]) {
        expect(num(sheet, row, year), cell(row, year)).toBe(0);
      }
    }
  });

  it.each<[string, Plan]>([
    ["the demo case", DEMO],
    ["a full sweep", { ...DEMO, sweep: 1 }],
    ["a cash shortfall", { ...DEMO, cfads: [20, 20, 20, 20, 20] }],
    [
      "a small loan swept out",
      { ...DEMO, loan: 30, sweep: 1, ebitda: [20, 22, 24, 26, 28] },
    ],
    [
      "a heavy amortiser at a high rate",
      { ...DEMO, loan: 400, rate: 0.11, amortisation: 0.15, sweep: 0.75 },
    ],
  ])("agrees with a second implementation on %s", (_name, plan) => {
    expectSchedule(evaluateTemplate(block, overrides(plan)), plan);
  });

  it("never lets a balance go negative or grow, whatever the plan, over 200 seeded plans", () => {
    const random = seeded(20_261_003);
    const between = (low: number, high: number): number =>
      low + (high - low) * random();
    for (let run = 0; run < 200; run += 1) {
      const plan: Plan = {
        loan: between(0, 600),
        rate: between(0, 0.15),
        amortisation: between(0, 0.25),
        sweep: between(0, 1),
        cfads: years.map(() => between(-60, 180)),
        ebitda: years.map(() => between(-20, 150)),
      };
      const sheet = evaluateTemplate(block, overrides(plan));
      expectSchedule(sheet, plan, `run ${String(run)}: `);
      for (const year of years) {
        const where = `run ${String(run)}, year ${String(year)}`;
        const closing = num(sheet, ROW.closing, year);
        expect(closing, where).toBeGreaterThanOrEqual(-1e-9);
        expect(closing, where).toBeLessThanOrEqual(
          num(sheet, ROW.opening, year) + 1e-9,
        );
        expect(num(sheet, ROW.sweep, year), where).toBeGreaterThanOrEqual(0);
        const ebitda = plan.ebitda[year - 1] as number;
        const leverage = sheet.value(cell(ROW.leverage, year));
        if (ebitda > 0) {
          expect(leverage as number, where).toBeCloseTo(closing / ebitda, 9);
        } else {
          expect(leverage, where).toBe("n.m.");
        }
      }
    }
  });

  it("shows an error as an error: a text rate is #VALUE!, nothing is guarded", () => {
    for (const row of block.cells) {
      for (const found of row) expect(found.f ?? "").not.toContain("IFERROR");
    }
    const sheet = evaluateTemplate(block, { [INPUT.rate]: "7 %" });
    expect(sheet.value(cell(ROW.interest, 1))).toEqual(
      new ExcelError("#VALUE!"),
    );
  });

  it("writes its formulas where a banker audits them", () => {
    const formula = (row: number, year: number): string | undefined => {
      const found = block.cells[row]?.[year];
      return found?.f === undefined ? undefined : resolveFormula(found.f, C5);
    };
    expect(formula(10, 1)).toBe("=D6");
    expect(formula(10, 2)).toBe("=D20");
    expect(formula(11, 1)).toBe("=D15*D7");
    expect(formula(12, 2)).toBe("=MIN(D8*D6,E15)");
    expect(formula(13, 1)).toBe("=D12-D16-D17");
    expect(formula(14, 1)).toBe("=MIN(D9*MAX(0,D18),D15-D17)");
    expect(formula(15, 1)).toBe("=D15-D17-D19");
    expect(formula(17, 1)).toBe('=IF(N(D13)>0,D20/D13,"n.m.")');
    expect(formula(17, 5)).toBe('=IF(N(H13)>0,H20/H13,"n.m.")');
  });
});
