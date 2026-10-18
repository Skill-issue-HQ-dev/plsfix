// Hunt 3: groups of TYPED decimals, the numbers a modeller really enters, held
// against an oracle that never touches a double. Every value is an integer
// count of 10^-k, so the total, the floors, the remainders and the ties are
// integer arithmetic (BigInt), and what the allocator should answer is not a
// matter of opinion: half goes away from zero, equal remainders are tied by
// position, the parts add up to the rounded total.
//
// The generator leans on exact halves and equal remainders (tails that are
// whole tenths of a unit), the groups the old allocator got wrong because it
// asked the binary double: 0.7 + 0.2 + 0.005 adds up to 0.9049999999999999.
// One filter keeps the oracle honest about Excel: when the doubles' own sum
// is visibly off the typed total at Excel's 15 digits (a big positive and a
// big negative cancelling leave noise a 15-digit cell shows), Excel itself
// rounds that noise, and so does pls,fix; there is no typed answer to hold
// it to, so the group is skipped.
import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { allocateRounded, roundedTotal } from "../../src/rounding";

const SEED = 20261007;
const RUNS = 3000;

// A typed number: `count` x 10^-places, kept apart from the double it becomes.
interface Typed {
  count: bigint;
}

interface Group {
  places: number; // k: decimals typed into every cell
  decimals: number; // d: decimals the group is rounded to, always below k
  cells: Typed[];
}

const pow10 = (power: number): bigint => 10n ** BigInt(power);

// One cell: a whole number of units plus a tail of the unit, signed. Tails of
// whole tenths of a unit and the exact half make equal remainders and exact
// half totals common instead of lucky.
function cellArb(unit: bigint): fc.Arbitrary<Typed> {
  const tail = fc.oneof(
    fc.constant(0n),
    fc.constant(unit / 2n),
    fc.integer({ min: 0, max: 9 }).map((tenth) => (BigInt(tenth) * unit) / 10n),
    fc.bigInt({ min: 0n, max: unit - 1n }),
  );
  return fc
    .tuple(fc.boolean(), fc.integer({ min: 0, max: 9999 }), tail)
    .map(([negative, units, rest]) => {
      const count = BigInt(units) * unit + rest;
      return { count: negative ? -count : count };
    });
}

const groupArb: fc.Arbitrary<Group> = fc
  .integer({ min: 1, max: 6 })
  .chain((places) =>
    fc.integer({ min: -3, max: places - 1 }).map((decimals) => ({
      places,
      decimals,
    })),
  )
  .chain(({ places, decimals }) => {
    const unit = pow10(places - decimals);
    const some = fc.array(cellArb(unit), { minLength: 1, maxLength: 30 });
    const same = fc
      .tuple(cellArb(unit), fc.integer({ min: 2, max: 25 }))
      .map(([cell, count]) => Array.from({ length: count }, () => cell));
    return fc
      .oneof({ weight: 3, arbitrary: some }, { weight: 1, arbitrary: same })
      .map((cells) => ({ places, decimals, cells }));
  });

// The number a cell holds: the double nearest the typed decimal.
const doubleOf = (cell: Typed, places: number): number =>
  Number(`${cell.count}e-${places}`);

// Half goes away from zero; a negative count floors down. All in BigInt.
function expected(group: Group): { total: number; parts: number[] } {
  const unit = pow10(group.places - group.decimals);
  const asNumber = (units: bigint): number =>
    Number(`${units}e${-group.decimals}`);

  const sum = group.cells.reduce((total, cell) => total + cell.count, 0n);
  const magnitude = sum < 0n ? -sum : sum;
  const roundedUp =
    magnitude / unit + ((magnitude % unit) * 2n >= unit ? 1n : 0n);
  const totalUnits = sum < 0n ? -roundedUp : roundedUp;

  const floors = group.cells.map((cell) => {
    const down = cell.count / unit;
    return cell.count % unit < 0n ? down - 1n : down;
  });
  const lost = group.cells.map(
    (cell, index) => cell.count - floors[index]! * unit,
  );
  const owed = totalUnits - floors.reduce((total, floor) => total + floor, 0n);
  // Exact arithmetic never owes a cell more than one unit, or less than none.
  expect(owed >= 0n && owed <= BigInt(group.cells.length)).toBe(true);

  // Largest remainder first, equal remainders in position order.
  const order = group.cells
    .map((_cell, index) => index)
    .sort((a, b) => {
      const first = lost[a]!;
      const second = lost[b]!;
      return first === second ? a - b : first > second ? -1 : 1;
    });
  const bumped = new Set(order.slice(0, Number(owed)));
  return {
    total: asNumber(totalUnits),
    parts: floors.map((floor, index) =>
      asNumber(floor + (bumped.has(index) ? 1n : 0n)),
    ),
  };
}

describe("groups of typed decimals", () => {
  it("round to the exact typed answer, ties by position, parts adding up to the total", () => {
    fc.assert(
      fc.property(groupArb, (group) => {
        const values = group.cells.map((cell) => doubleOf(cell, group.places));
        // The skip the header names: only a total Excel reads as typed.
        const typedTotal = Number(
          `${group.cells.reduce((total, cell) => total + cell.count, 0n)}e-${group.places}`,
        );
        const sum = values.reduce((total, value) => total + value, 0);
        fc.pre(Number(sum.toPrecision(15)) === typedTotal);

        const want = expected(group);
        expect(roundedTotal(values, group.decimals)).toBe(want.total);
        expect(allocateRounded(values, group.decimals)).toEqual(want.parts);
      }),
      { seed: SEED, numRuns: RUNS },
    );
  });
});
