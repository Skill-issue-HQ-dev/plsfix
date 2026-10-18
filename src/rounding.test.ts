// The rounding contract: the parts add up to the total at the same precision,
// every part stays within one unit of its value, and the same inputs always
// produce the same allocation - the reason a hundred sibling cells can each
// recompute the group on their own and still agree.
import { describe, expect, it } from "vitest";
import {
  allocateRounded,
  ROUNDING_MAX_DECIMALS,
  roundedTotal,
} from "./rounding";

// Adding the parts back up in the same precision the caller asked for: the
// invariant the whole feature exists to keep.
function partsTotal(parts: number[], decimals: number): number {
  return roundedTotal(parts, decimals);
}

describe("allocateRounded", () => {
  it("splits 100 across three thirds without losing the unit", () => {
    const values = [33.333, 33.333, 33.334];
    expect(allocateRounded(values, 0)).toEqual([33, 33, 34]);
    expect(partsTotal(allocateRounded(values, 0), 0)).toBe(
      roundedTotal(values, 0),
    );
    expect(roundedTotal(values, 0)).toBe(100);
  });

  it("breaks a tie by position, not by value", () => {
    const third = 10 / 3;
    expect(allocateRounded([third, third, third], 0)).toEqual([4, 3, 3]);
    expect(allocateRounded([third, third, third], 1)).toEqual([3.4, 3.3, 3.3]);
  });

  it("mirrors the split for negative values", () => {
    const values = [-33.333, -33.333, -33.334];
    expect(allocateRounded(values, 0)).toEqual([-33, -33, -34]);
    expect(roundedTotal(values, 0)).toBe(-100);
  });

  it("allocates across mixed signs", () => {
    const values = [2.5, -1.2, -0.4];
    expect(allocateRounded(values, 0)).toEqual([2, -1, 0]);
    expect(roundedTotal(values, 0)).toBe(1);
  });

  it("leaves zeros and whole numbers alone", () => {
    expect(allocateRounded([0, 0, 0], 0)).toEqual([0, 0, 0]);
    expect(allocateRounded([10, 20, 30], 2)).toEqual([10, 20, 30]);
    expect(allocateRounded([], 0)).toEqual([]);
  });

  it("rounds a single value half away from zero", () => {
    expect(allocateRounded([2.5], 0)).toEqual([3]);
    expect(allocateRounded([-2.5], 0)).toEqual([-3]);
    expect(roundedTotal([-0.5], 0)).toBe(-1);
    expect(roundedTotal([0.5], 0)).toBe(1);
  });

  it("keeps decimals exact rather than binary-approximate", () => {
    // 1.005 * 100 is 100.49999999999999 in binary; a multiply here would floor
    // to 100 and hand the unit to the wrong cell.
    const values = [1.005, 2.005, 3.005];
    expect(allocateRounded(values, 2)).toEqual([1.01, 2.01, 3]);
    expect(roundedTotal(values, 2)).toBe(6.02);
  });

  it("rounds to thousands on negative decimals", () => {
    expect(allocateRounded([1400, 1400, 1400], -3)).toEqual([2000, 1000, 1000]);
    expect(roundedTotal([1400, 1400, 1400], -3)).toBe(4000);
  });

  // A hand-checked block: the parts a modeller would read off the sheet.
  it("keeps a percentage split adding to the whole", () => {
    const values = [0.31428, 0.31428, 0.37144];
    expect(allocateRounded(values, 3)).toEqual([0.314, 0.314, 0.372]);
    expect(roundedTotal(values, 3)).toBe(1);
  });

  it("holds the total and the one-unit bound over a long series", () => {
    // Deterministic pseudo-random values, so a failure is reproducible.
    let seed = 7;
    const values = Array.from({ length: 200 }, () => {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return (seed / 2147483648) * 2000 - 500;
    });

    for (const decimals of [-1, 0, 1, 2]) {
      const parts = allocateRounded(values, decimals);
      expect(partsTotal(parts, decimals)).toBe(roundedTotal(values, decimals));
      const unit = 10 ** -decimals;
      parts.forEach((part, index) => {
        expect(Math.abs(part - values[index]!)).toBeLessThan(unit);
      });
    }
  });

  it("gives every caller the same allocation", () => {
    const values = [1.4, 1.4, 1.4, 1.4, 1.4];
    const first = allocateRounded(values, 0);
    expect(allocateRounded([...values], 0)).toEqual(first);
    expect(first).toEqual([2, 2, 1, 1, 1]);
  });

  // Excel reads a number at 15 significant digits before it rounds, so the
  // total of 0.7 + 0.2 + 0.005 (0.9049999999999999 as a double) is the
  // half 0.905 to it, and a half goes away from zero on either side of zero.
  it("rounds an exact-half total the way Excel does", () => {
    expect(roundedTotal([0.7, 0.2, 0.005], 2)).toBe(0.91);
    expect(allocateRounded([0.7, 0.2, 0.005], 2)).toEqual([0.7, 0.2, 0.01]);
    expect(roundedTotal([-0.7, -0.2, -0.005], 2)).toBe(-0.91);
    expect(allocateRounded([-0.7, -0.2, -0.005], 2)).toEqual([
      -0.7, -0.2, -0.01,
    ]);
  });

  it("ties equal remainders by position whatever the size of the values", () => {
    // Each of these loses exactly .33 of a unit to the floor; as doubles the
    // remainders of 3000.33 and 20.33 differ from the 14th decimal on.
    expect(allocateRounded([3000.33, 20.33, 0.33], 0)).toEqual([3001, 20, 0]);
    expect(allocateRounded([0.33, 20.33, 3000.33], 0)).toEqual([1, 20, 3000]);
    // Negative halves tie the same way: -1000.5, -0.5 and -20.5 all lose .5.
    expect(allocateRounded([-1000.5, -0.5, -20.5], 0)).toEqual([
      -1000, -1, -21,
    ]);
  });

  it("reads a number at 15 digits even where the rounding unit lies beyond them", () => {
    // 1447182377411.3337 has 17 digits; Excel holds 1447182377411.33, the 3rd
    // decimal is past its 15th digit, so rounding to 3 decimals keeps .33.
    expect(roundedTotal([1447182377411.3337], 3)).toBe(1447182377411.33);
    expect(allocateRounded([1447182377411.3337], 3)).toEqual([
      1447182377411.33,
    ]);
  });

  it("scales a value past 1e308 in units without overflowing", () => {
    expect(roundedTotal([1e300], 10)).toBe(1e300);
    expect(allocateRounded([1e300, 1], 10)).toEqual([1e300, 1]);
    // A tiny value is the other way: its units are far below 1, and it is
    // the smallest remainder in the group, never the one that is bumped.
    expect(allocateRounded([5e-324, 1e-300, 0.5], 2)).toEqual([0, 0, 0.5]);
  });

  it("refuses a total no double can hold, and returns an answer that large as Infinity", () => {
    expect(() => roundedTotal([1e308, 1e308], 0)).toThrow(
      "rounding: these values add up to more than Excel can hold",
    );
    expect(() => allocateRounded([1e308, 1e308], 0)).toThrow(
      "rounding: these values add up to more than Excel can hold",
    );
    // 15 digits of the largest double round up past it: for the caller to refuse.
    expect(roundedTotal([Number.MAX_VALUE], 0)).toBe(Infinity);
    expect(allocateRounded([Number.MAX_VALUE], 0)).toEqual([Infinity]);
  });

  it("refuses values and precisions it cannot allocate", () => {
    expect(() => allocateRounded([1, Number.NaN], 0)).toThrow(
      "rounding: every value must be a finite number",
    );
    expect(() => allocateRounded([1, Number.POSITIVE_INFINITY], 0)).toThrow(
      "rounding: every value must be a finite number",
    );
    expect(() => allocateRounded([1, 2], 1.5)).toThrow(
      "rounding: decimals must be a whole number",
    );
    expect(() => allocateRounded([1, 2], ROUNDING_MAX_DECIMALS + 1)).toThrow(
      "rounding: decimals must be a whole number",
    );
    expect(() => roundedTotal([1, 2], Number.NaN)).toThrow(
      "rounding: decimals must be a whole number",
    );
  });
});
