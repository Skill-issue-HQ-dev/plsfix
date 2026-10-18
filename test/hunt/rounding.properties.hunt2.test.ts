// Pass-2 property: the largest-remainder allocator's guarantees, hammered
// with fast-check across decimals -10..10 and values spanning 1e-9..1e12 of
// both signs, zeros and repeated values. They are exact guarantees, with no
// tolerance standing in for binary noise: the group's total is
// =ROUND(SUM(range), decimals) as test/formula-functions.ts computes Excel's
// answer, and the parts add up to it in whole units.
//
// "Exact" is a claim about numbers Excel itself can hold. Excel keeps 15
// significant digits, so a value is only as exact as that: every group here is
// kept under a number of whole rounding units that leaves digits below the
// unit (src/rounding.test.ts, "reads a number at 15 digits even where the
// rounding unit lies beyond them", pins what the allocator does past them, and
// why that is not a bug to chase).
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import {
  allocateRounded,
  ROUNDING_MAX_DECIMALS,
  roundedTotal,
} from "../../src/rounding";
import { excelRound } from "../formula-functions";

const SEED = 20260927;
const RUNS = 3000;

// Sign, mantissa in [1,10) and a power-of-ten exponent from -9 to 12: the
// exact magnitude spread the brief names, not a single fc.double() range
// that would rarely sample the small end.
const magnitudeValueArb = fc
  .tuple(
    fc.boolean(),
    fc.double({ min: 1, max: 10, noNaN: true }),
    fc.integer({ min: -9, max: 12 }),
  )
  .map(([negative, mantissa, exponent]) => {
    const value = mantissa * 10 ** exponent;
    return negative ? -value : value;
  });

const valueArb = fc.oneof(
  { weight: 1, arbitrary: fc.constant(0) },
  { weight: 4, arbitrary: magnitudeValueArb },
);

const decimalsArb = fc.integer({
  min: -ROUNDING_MAX_DECIMALS,
  max: ROUNDING_MAX_DECIMALS,
});

// A plain random group, or the same value repeated several times - the
// brief's "repeated values" case, which is also the group most likely to
// expose a tie-break bug (every remainder equal).
const valuesArb = fc.oneof(
  { weight: 3, arbitrary: fc.array(valueArb, { minLength: 0, maxLength: 30 }) },
  {
    weight: 1,
    arbitrary: fc
      .tuple(valueArb, fc.integer({ min: 2, max: 25 }))
      .map(([value, count]) => Array.from({ length: count }, () => value)),
  },
);

// A value in whole units of 10^-decimals: the unit count of anything a part or
// a total can be, so sums and comparisons below are integer arithmetic. The
// product sits within 1e-4 of the integer it stands for, so rounding it back
// is exact rather than a tolerance.
const wholeUnits = (value: number, decimals: number): number =>
  Math.round(decimals >= 0 ? value * 10 ** decimals : value / 10 ** -decimals);

// Groups whose every value is under `maxUnits` whole units at the decimals.
function casesUnder(maxUnits: number) {
  return fc
    .tuple(valuesArb, decimalsArb)
    .filter(([values, decimals]) =>
      values.every((value) => Math.abs(value) * 10 ** decimals < maxUnits),
    );
}

// Four digits of room under the unit: the total below is ROUND(SUM()) with
// no difference in millions of random groups, and the parts always add up.
// The bound is the evaluator's, not the allocator's: excelRound cuts the SCALED
// total to 15 digits, a double product with a rounding of its own, where
// src/rounding.ts cuts the total itself (what the cell shows) and scales it
// exactly. Past about 1e12 units the two read a sum's 15th digit differently in
// a few groups in a hundred thousand, a unit apart.
const exactCaseArb = casesUnder(1e11);

// The widest groups of all, a double's exact-integer range, for the one
// property that does not lean on Excel's digits.
const wideCaseArb = casesUnder(2 ** 51);

describe("the rounded total is =ROUND(SUM(range), decimals)", () => {
  it("equals the Excel evaluator's answer exactly, for every generated group", () => {
    fc.assert(
      fc.property(exactCaseArb, ([values, decimals]) => {
        // Excel's SUM adds the doubles in order; its ROUND is excelRound.
        const sum = values.reduce((total, value) => total + value, 0);
        expect(roundedTotal(values, decimals)).toBe(excelRound(sum, decimals));
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("allocateRounded sums exactly to roundedTotal", () => {
  it("holds, in whole units, over generated groups at every decimals in range", () => {
    fc.assert(
      fc.property(exactCaseArb, ([values, decimals]) => {
        const parts = allocateRounded(values, decimals);
        const sum = parts.reduce(
          (total, part) => total + wholeUnits(part, decimals),
          0,
        );
        expect(sum).toBe(wholeUnits(roundedTotal(values, decimals), decimals));
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("every part is within one unit of its own value", () => {
  it("never lands a part more than one unit from what Excel rounds the value to", () => {
    fc.assert(
      fc.property(exactCaseArb, ([values, decimals]) => {
        const parts = allocateRounded(values, decimals);
        parts.forEach((part, index) => {
          // A part is its value floored or one unit up, and ROUND is one of
          // those two, so they are never further apart than one unit.
          const nearest = excelRound(values[index]!, decimals);
          expect(
            Math.abs(
              wholeUnits(part, decimals) - wholeUnits(nearest, decimals),
            ),
          ).toBeLessThanOrEqual(1);
        });
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

// bumpOrder (src/rounding.ts) breaks a tied remainder by POSITION IN WHATEVER
// ARRAY IT IS GIVEN, by design ("equal remainders keep their position order,
// so the allocation is the same for every sibling cell that recomputes it" -
// every sibling reads the SAME range in the SAME order). It does not, and is
// not meant to, track a value's identity through an arbitrary reordering: two
// tied entries that trade places can legitimately trade which one gets the
// bump. So this property is scoped to groups with no tied remainder, where
// there is nothing left for a permutation to disturb.
describe("permuting the input permutes the output the same way", () => {
  function remainder(value: number, decimals: number): number {
    const scaled = value * 10 ** decimals;
    return scaled - Math.floor(scaled);
  }

  // Excel reads each value at 15 digits, which moves a remainder by up to
  // 5e-15 of the value in units; under 1e8 units that is below the 1e-6 gap
  // kept here, so two remainders this far apart are not tied once read.
  function hasCloseRemainders(values: number[], decimals: number): boolean {
    const remainders = values.map((value) => remainder(value, decimals));
    for (let i = 0; i < remainders.length; i += 1) {
      for (let j = i + 1; j < remainders.length; j += 1) {
        if (Math.abs(remainders[i]! - remainders[j]!) < 1e-6) return true;
      }
    }
    return false;
  }

  it("moves each part with its value under any reordering, no remainder tied", () => {
    fc.assert(
      fc.property(
        casesUnder(1e8)
          .filter(([values]) => values.length >= 2)
          .filter(([values, decimals]) => !hasCloseRemainders(values, decimals))
          .chain(([values, decimals]) => {
            const indices = values.map((_unused, index) => index);
            return fc.tuple(
              fc.constant(values),
              fc.constant(decimals),
              fc.shuffledSubarray(indices, {
                minLength: indices.length,
                maxLength: indices.length,
              }),
            );
          }),
        ([values, decimals, permutation]) => {
          const permutedValues = permutation.map((index) => values[index]!);
          // Floating-point addition is not associative: summing the same
          // multiset in a different order can land on a different double
          // when magnitudes are far enough apart (values here can span
          // 1e-9..1e12), and the total is cut from that double sum. That is
          // a property of IEEE754 summation order, not of this allocator's
          // tie-breaking, so it is out of scope here - skip the rare
          // instance where this permutation actually changes the float sum,
          // rather than let it masquerade as a tie-break bug.
          const add = (total: number, value: number) => total + value;
          if (values.reduce(add, 0) !== permutedValues.reduce(add, 0)) return;

          const original = allocateRounded(values, decimals);
          const permutedResult = allocateRounded(permutedValues, decimals);
          const expected = permutation.map((index) => original[index]!);
          expect(permutedResult).toEqual(expected);
        },
      ),
      { seed: SEED, numRuns: RUNS },
    );
  });
});

describe("no -0 anywhere in the output", () => {
  it("never hands back a negative zero, part or total", () => {
    fc.assert(
      fc.property(wideCaseArb, ([values, decimals]) => {
        const parts = allocateRounded(values, decimals);
        for (const part of parts) expect(Object.is(part, -0)).toBe(false);
        expect(Object.is(roundedTotal(values, decimals), -0)).toBe(false);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
