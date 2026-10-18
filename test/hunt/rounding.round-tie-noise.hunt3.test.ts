// Hunt r1 / H5 (fixed in wave FD): =PLSFIX.ROUND hands the units still owed to "the
// parts that lost the most, ties by position" (src/rounding.ts header, and the doc
// comment of `bumpOrder`: "equal remainders keep their position order"). The
// remainder was `unit - floor` on doubles, so two rows that lost exactly the same
// decimal amount (.33 each) compared unequal by binary noise that depends on how big
// the row's integer part is, and the extra unit went to whichever row the noise
// favoured, not to the first one.
// Why the expectation is right: the decimal values 3000.33, 20.33 and 0.33 all lose
// exactly 0.33 of a unit to the floor, and the module says ties are broken by position.
import { afterEach, describe, expect, it, vi } from "vitest";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

type Round = (range: number[][], index: number, decimals: number) => number;

async function load(): Promise<Round> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => undefined,
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  return (await import("../../src/functions/index")).smtRound;
}

/** What the group's cells show: PLSFIX.ROUND for index 1..n over one range. */
function column(round: Round, values: number[], decimals: number): number[] {
  const range = values.map((value) => [value]);
  return values.map((_, at) => round(range, at + 1, decimals));
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("equal remainders are tied by position", () => {
  it("gives the one owed unit of 3000.33 + 20.33 + 0.33 to the FIRST row at 0 decimals", async () => {
    const round = await load();
    // 3000.33 + 20.33 + 0.33 = 3020.99 -> 3021; floors add to 3020; one unit owed;
    // every row lost exactly .33, so the first row takes it.
    expect(column(round, [3000.33, 20.33, 0.33], 0)).toEqual([3001, 20, 0]);
  });

  it("gives the three owed tenths of 1000.13 + nine times 0.13 to the first three rows", async () => {
    const round = await load();
    const values = [1000.13, ...Array.from({ length: 9 }, () => 0.13)];
    // Every row lost exactly 0.3 of a tenth; three tenths are owed; rows 1-3 take them.
    expect(column(round, values, 1)).toEqual([
      1000.2, 0.2, 0.2, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1,
    ]);
  });
});
