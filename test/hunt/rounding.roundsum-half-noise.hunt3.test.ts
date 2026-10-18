// Hunt r1 / H5 (fixed in wave FD): =PLSFIX.ROUNDSUM must round the group's DECIMAL
// total, the way the same total typed into one cell rounds. src/rounding.ts used to
// shift every single value through its decimal string ("1.005e2" = 100.5, exact) but
// to add the values as doubles first, so binary noise (0.7 + 0.2 + 0.005 =
// 0.9049999999999999) decided an exact half. Why the expectation is right: Excel's
// ROUND reads a number at 15 significant digits (0.905) and sends a half away from
// zero, and PLSFIX.ROUNDSUM([0.905], 2) was 0.91 in this very module; the decimal
// sum 0.7 + 0.2 + 0.005 is exactly 0.905.
import { afterEach, describe, expect, it, vi } from "vitest";

import type { Template } from "../../src/template-cells";
import { evaluateTemplate } from "../formula-eval";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

type RoundSum = (range: number[][], decimals: number) => number;
type Round = (range: number[][], index: number, decimals: number) => number;

async function load(): Promise<{ roundSum: RoundSum; round: Round }> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => undefined,
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  const module = await import("../../src/functions/index");
  return { roundSum: module.smtRoundSum, round: module.smtRound };
}

const column = (values: number[]): number[][] => values.map((v) => [v]);

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("ROUNDSUM at an exact half", () => {
  it("rounds 0.7 + 0.2 + 0.005 (decimal total 0.905) like the single cell 0.905", async () => {
    const { roundSum } = await load();
    // The reference: one cell holding the same decimal total.
    expect(roundSum(column([0.905]), 2)).toBe(0.91);
    // Same total, split in three: must be the same answer.
    expect(roundSum(column([0.7, 0.2, 0.005]), 2)).toBe(0.91);
  });

  it("gives two groups with the same exact-half total the same rounding direction", async () => {
    const { roundSum } = await load();
    // 0.1 + 0.2 + 0.005 = 0.305 and 0.3 + 0.6 + 0.005 = 0.905: both are .xx5 halves.
    const up = roundSum(column([0.1, 0.2, 0.005]), 2); // 0.31
    const sameKind = roundSum(column([0.3, 0.6, 0.005]), 2); // must be 0.91, is 0.9
    expect([up, sameKind]).toEqual([0.31, 0.91]);
  });

  it("does not depend on how a 0.035 total is split into two cells", async () => {
    const { roundSum } = await load();
    expect(roundSum(column([0.035]), 2)).toBe(0.04);
    expect(roundSum(column([0.005, 0.03]), 2)).toBe(0.04);
    expect(roundSum(column([0.03, 0.005]), 2)).toBe(0.04);
  });

  it("agrees with =ROUND(SUM(range), 2) as the repo's own Excel evaluator computes it", async () => {
    const { roundSum } = await load();
    // test/formula-eval.ts models Excel's ROUND (15 significant digits, then half away
    // from zero): the same cells through ROUND(SUM()) are what a modeller compares with.
    const groups = [
      [0.7, 0.2, 0.005],
      [0.3, 0.6, 0.005],
      [0.005, 0.03],
      [85.6, 1850, 13.08, 0.36, 14.21, 1.245],
    ];
    const rows = groups.map((values) => {
      const cells = [
        ...values.map((value) => [{ v: value }]),
        [{ f: `=ROUND(SUM(A1:A${values.length}),2)` }],
      ];
      const excel = evaluateTemplate({ cells } as unknown as Template).value(
        `A${values.length + 1}`,
      );
      return { cells: values, roundSum: roundSum(column(values), 2), excel };
    });
    expect(rows.map((row) => row.roundSum)).toEqual(
      rows.map((row) => row.excel),
    );
  });

  it("keeps the parts of PLSFIX.ROUND adding up to that half-rounded total", async () => {
    const { round } = await load();
    const range = column([0.7, 0.2, 0.005]);
    const parts = [1, 2, 3].map((index) => round(range, index, 2));
    const cents = Math.round(parts.reduce((sum, part) => sum + part, 0) * 100);
    expect(cents).toBe(91); // 0.91, the rounded decimal total
  });
});
