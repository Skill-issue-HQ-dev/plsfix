// The three-statement mini's check row against what it promises: zero in every
// column while the sheet balances and the net-debt memo ties to debt less cash,
// a gap the moment a balance or a link is typed over, and no way for two wrong
// figures to cancel. Computed by test/formula-eval.ts; the gaps are worked out
// by hand on the demo.
import { describe, expect, it } from "vitest";
import { evaluateTemplate, type Value } from "../../test/formula-eval";
import { threeStatement } from "./three-statement";

// With the block's corner at A1: the opening balances in B, years one to three
// in C to E. Rows: 6 D&A, 8 interest rate, 11 tax rate, 17 cash, 19 equity, 20 net
// debt memo, 21 the check, 23 the D&A add-back, 24 capex, 25 working capital,
// 26 free cash flow.
const COLUMNS = ["B", "C", "D", "E"] as const;
const block = threeStatement();

function checks(overrides: Record<string, Value>): unknown[] {
  const sheet = evaluateTemplate(block, overrides);
  return COLUMNS.map((column) => sheet.value(`${column}21`));
}

describe("the three-statement mini's check row", () => {
  it("is zero in the opening column and every year on the demo", () => {
    expect(checks({})).toEqual([0, 0, 0, 0]);
  });

  it("stays at zero when an input changes, because every flow moves both sides", () => {
    // Year-one capex, year-two D&A and tax rate, year-three interest rate and a
    // year-two working-capital increase typed over: each one is a different
    // business, not a broken link.
    const scenarios: Record<string, Value>[] = [
      { C24: 20 },
      { D6: 9 },
      { D11: 0.35 },
      { E8: 0.2 },
      { D25: -4 },
    ];
    for (const overrides of scenarios) {
      expect(checks(overrides), JSON.stringify(overrides)).toEqual([
        0, 0, 0, 0,
      ]);
    }
  });

  it("shows the gap when an opening balance is typed over, and keeps showing it", () => {
    // by hand: equity typed 45 against 60 + 20 + 10 - 50 = 40 of net assets is 5
    // too high; every year rolls from it, so the 5 stays.
    expect(checks({ B19: 45 })).toEqual([5, 5, 5, 5]);
  });

  it("shows the gap when year one's closing cash, the year-two opening cash, is typed over", () => {
    // by hand: cash typed 20 instead of 10 + 8 = 18 puts the sheet 62 + 22 + 20
    // - 50 - 52 = 2 out; year two rolls from the typed 20, so the 2 stays.
    expect(checks({ C17: 20 })).toEqual([0, 2, 2, 2]);
  });

  it("shows the gap when year two's closing net debt, the year-three opening net debt, is typed over", () => {
    // by hand: year two's net debt is 50 - 28.056 = 21.944; typed 20 the memo is
    // 1.944 off its own debt less cash. Year three opens on the typed 20, which
    // moves its interest, but year three's memo is a formula again: 0.
    const got = checks({ D20: 20 });
    expect(got.slice(0, 2)).toEqual([0, 0]);
    expect(got[2]).toBeCloseTo(1.944, 9);
    expect(got[3]).toBe(0);
  });

  it("shows the gap when year two's free cash flow is typed over", () => {
    // by hand: FCF should be 10.056; typed 12 it adds 1.944 to cash that no
    // other line records, so assets run 1.944 ahead of debt and equity, and
    // year three rolls from the wrong cash.
    const got = checks({ D26: 12 });
    expect(got.slice(0, 2)).toEqual([0, 0]);
    expect(got[2]).toBeCloseTo(1.944, 9);
    expect(got[3]).toBeCloseTo(1.944, 9);
  });

  it("shows the gap when year two's D&A add-back is typed over, though the D&A input is not", () => {
    // by hand: D&A of 6.5 comes back into free cash flow; typed 5 in the add-back
    // takes 1.5 off cash while PP&E still falls by 6.5, so the sheet is 1.5 out.
    const got = checks({ D23: 5 });
    expect(got.slice(0, 2)).toEqual([0, 0]);
    expect(got[2]).toBeCloseTo(1.5, 9);
    expect(got[3]).toBeCloseTo(1.5, 9);
  });

  it("does not let two wrong figures cancel: a balance gap and a memo gap each count", () => {
    // by hand: cash typed 20 is 2 too high (sheet +2); net debt typed 28 where
    // 50 - 20 = 30 is 2 too low (memo -2). Summed they would read 0, taken at
    // their distance they read 2 + 2 = 4.
    expect(checks({ C17: 20, C20: 28 })[1]).toBe(4);
  });
});
