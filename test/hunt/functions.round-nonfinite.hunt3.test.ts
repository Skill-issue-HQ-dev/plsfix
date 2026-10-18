// Hunt r1 / H5 (fixed in wave FD): a custom function must answer a number Excel can
// print or a #VALUE! with a sentence, never NaN. =PLSFIX.CAGR already refuses a result
// that is not finite (src/functions/index.ts, "A refusal is the only honest answer"),
// but ROUND and ROUNDSUM only checked their ARGUMENTS: two finite cells whose sum
// overflows a double, or a value that overflows once it is scaled by 10^decimals,
// came back as NaN.
// Why the expectation is right: the same file says a number Excel cannot print is not
// an answer (smtCagr), and test/hunt/functions.properties.hunt2 holds CAGR to "a number
// CustomFunctions can actually print, never Infinity"; Excel's own SUM(1E308, 1E308)
// is #NUM!, not a number.
import { afterEach, describe, expect, it, vi } from "vitest";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

interface Loaded {
  round: (range: number[][], index: number, decimals: number) => number;
  roundSum: (range: number[][], decimals: number) => number;
}

async function load(): Promise<Loaded> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => undefined,
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  const module = await import("../../src/functions/index");
  return { round: module.smtRound, roundSum: module.smtRoundSum };
}

/** What the cell would hold: the number, or the error code the function raised. */
function answer(run: () => number): number | string {
  try {
    return run();
  } catch (error) {
    return error instanceof FakeFunctionError ? error.code : String(error);
  }
}

/** The two honest answers: a finite number, or the documented #VALUE! refusal. */
function printable(got: number | string): boolean {
  return typeof got === "number" ? Number.isFinite(got) : got === "#VALUE!";
}

const column = (values: number[]): number[][] => values.map((v) => [v]);

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("ROUND and ROUNDSUM never answer NaN for finite arguments", () => {
  it("ROUNDSUM over two cells of 1E+308 (the sum overflows)", async () => {
    const { roundSum } = await load();
    const got = answer(() => roundSum(column([1e308, 1e308]), 0));
    expect(printable(got), `the cell holds ${String(got)}`).toBe(true);
  });

  it("ROUNDSUM of 1E+300 at 10 decimals (the scaled value overflows)", async () => {
    const { roundSum } = await load();
    const got = answer(() => roundSum(column([1e300]), 10));
    expect(printable(got), `the cell holds ${String(got)}`).toBe(true);
  });

  it("ROUND of the first cell of {1E+300, 1} at 10 decimals", async () => {
    const { round } = await load();
    const got = answer(() => round(column([1e300, 1]), 1, 10));
    expect(printable(got), `the cell holds ${String(got)}`).toBe(true);
  });
});
