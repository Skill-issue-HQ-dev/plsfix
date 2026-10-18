// Hunt r1 / H5 (fixed in wave FD): =PLSFIX.CAGR(first, last, periods) with last /
// first so small that the ratio underflows to 0 answered exactly -100%, which is the
// rate of an END VALUE OF ZERO, while the function refuses a zero end value in the
// same breath ("a positive end value"). src/chartmath.ts `cagr` rescued the OVERFLOW
// side in log space ("The ratio overflowed a double although a rate may still exist")
// and left the mirror image alone: `Number.isFinite(0)` is true, so
// `0 ** (1 / periods) - 1` = -1.
// Why the expectation is right: for last > 0 the true rate is exp((ln last - ln first)
// / periods) - 1, strictly above -1 for any finite number of periods; -0.7488 here.
import { afterEach, describe, expect, it, vi } from "vitest";

class FakeFunctionError {
  constructor(
    public code: string,
    public message?: string,
  ) {}
}

async function loadCagr(): Promise<
  (first: number, last: number, periods: number) => number
> {
  vi.stubGlobal("CustomFunctions", {
    associate: () => undefined,
    Error: FakeFunctionError,
    ErrorCode: { invalidValue: "#VALUE!" },
  });
  vi.resetModules();
  return (await import("../../src/functions/index")).smtCagr;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("CAGR where the end is far below the start", () => {
  it("is not -100% for a positive end value (1E+300 down to 1E-300 over 1000 periods)", async () => {
    const cagr = await loadCagr();
    const rate = cagr(1e300, 1e-300, 1000);
    const expected = Math.exp((Math.log(1e-300) - Math.log(1e300)) / 1000) - 1;
    expect(rate).toBeGreaterThan(-1);
    expect(rate).toBeCloseTo(expected, 9); // -0.74881...
  });

  it("agrees with the log-space answer one decade inside the overflow edge (1E+200 to 1E-200 over 100 periods)", async () => {
    const cagr = await loadCagr();
    const rate = cagr(1e200, 1e-200, 100);
    const expected = Math.exp((Math.log(1e-200) - Math.log(1e200)) / 100) - 1;
    // exp(-9.21) - 1 = -0.9999, not -1
    expect(rate).toBeCloseTo(expected, 9);
    expect(rate).toBeGreaterThan(-1);
  });
});
