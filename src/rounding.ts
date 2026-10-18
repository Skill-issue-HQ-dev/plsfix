// Consistent rounding, think-cell TCROUND style: the rounded parts of a group
// add up to the rounded total, instead of missing it by a unit. Largest
// remainder - every part is floored, and the units still owed to the total go
// to the parts that lost the most, ties by position.
//
// The arithmetic is Excel's, not the double's: Excel holds 15 significant
// digits of every number, so 0.9049999999999999 (what 0.7 + 0.2 + 0.005 adds
// up to) is 0.905 to it and ROUND(.., 2) answers 0.91. Every value is cut to
// those 15 digits and kept as an exact decimal, so units, remainders and ties
// are whole-number sums with no binary noise in them.
//
// Pure by design and imports nothing: the custom-functions bundle
// (src/functions/) loads this module into Excel's function runtime, where none
// of the pane's DOM or Office.js code may follow it.

// Above this many cells the group stops being free to recalculate: every
// sibling PLSFIX.ROUND cell redoes the whole allocation, so one edit costs
// O(N^2 log N). Reconciliation blocks - the case this is written for - are
// dozens of rows, not thousands (docs/research/custom-functions.md, Risks).
export const ROUNDING_CELL_CAP = 1_000;

// Where the decimals argument ends, the range the functions document and
// refuse outside of. The arithmetic below is exact at any precision.
export const ROUNDING_MAX_DECIMALS = 10;

// What Excel keeps of a number, and shows of it: 15 significant digits.
const EXCEL_DIGITS = 15;

// A number as Excel holds it: the exact decimal `digits` x 10^`scale`, with at
// most 15 significant digits in `digits`.
interface Held {
  digits: bigint;
  scale: number;
}

// A value in whole units of 10^-decimals, floored towards minus infinity, and
// what the floor dropped as the exact fraction `dropped` / 10^`places`.
interface Units {
  floor: bigint;
  dropped: bigint;
  places: number;
}

const tenPower = (power: number): bigint => 10n ** BigInt(power);

// The 15-digit cut Excel makes of a double, as an exact decimal. toExponential
// rounds the double's own binary value to the nearest 15-digit decimal (a tie
// goes up in size), the same cut as Number(x.toPrecision(15)) without going
// back through a double again: 1.005 is the decimal 1.005 here, not 1.00499...
function held(value: number): Held {
  const [mantissa = "0", exponent = "0"] = value
    .toExponential(EXCEL_DIGITS - 1)
    .split("e");
  return {
    digits: BigInt(mantissa.replace(".", "")),
    scale: Number(exponent) - (EXCEL_DIGITS - 1),
  };
}

// Excel's own ROUND halves go away from zero, so -0.5 rounds to -1, not to 0.
// `held` is read in whole units of 10^-decimals: the unit count of the rounded
// number, decided on exact digits so an exact half is never a near miss.
function unitsHalfAway(number: Held, decimals: number): bigint {
  const at = number.scale + decimals;
  if (at >= 0) return number.digits * tenPower(at);
  const base = tenPower(-at);
  const magnitude = number.digits < 0n ? -number.digits : number.digits;
  // The dropped part is half a unit or more: the magnitude goes up one.
  const up = (magnitude % base) * 2n >= base ? 1n : 0n;
  const units = magnitude / base + up;
  return number.digits < 0n ? -units : units;
}

// The number as whole units of 10^-decimals, floored, with the exact rest.
function unitsFloored(number: Held, decimals: number): Units {
  const at = number.scale + decimals;
  if (at >= 0 || number.digits === 0n) {
    return {
      floor: number.digits * tenPower(Math.max(at, 0)),
      dropped: 0n,
      places: 0,
    };
  }
  const base = tenPower(-at);
  // BigInt division truncates towards zero and the remainder keeps the sign of
  // the dividend, so a negative value's floor is one unit lower.
  let floor = number.digits / base;
  let dropped = number.digits % base;
  if (dropped < 0n) {
    floor -= 1n;
    dropped += base;
  }
  return { floor, dropped, places: -at };
}

function assertInputs(values: number[], decimals: number): void {
  if (
    !Number.isInteger(decimals) ||
    Math.abs(decimals) > ROUNDING_MAX_DECIMALS
  ) {
    throw new Error(
      `rounding: decimals must be a whole number between -${ROUNDING_MAX_DECIMALS} and ${ROUNDING_MAX_DECIMALS}`,
    );
  }
  if (values.some((value) => !Number.isFinite(value))) {
    throw new Error("rounding: every value must be a finite number");
  }
}

// The grand total in whole units of 10^-decimals: what the parts have to add
// up to, and the one place the target is decided for both exports. It is
// ROUND(SUM(range), decimals) as Excel computes it: the doubles are added as
// SUM adds them, in order, and only then is the total cut to 15 digits.
function totalUnits(values: number[], decimals: number): bigint {
  const total = values.reduce((sum, value) => sum + value, 0);
  if (!Number.isFinite(total)) {
    throw new Error(
      "rounding: these values add up to more than Excel can hold",
    );
  }
  return unitsHalfAway(held(total), decimals);
}

// Whole units of 10^-decimals back to the number a cell shows. The decimal
// string parses to the double nearest the exact value, and a zero is never
// minus zero. Past the largest double the answer is Infinity, for the caller
// to refuse (src/functions/index.ts).
function fromUnits(units: bigint, decimals: number): number {
  return Number(`${units}e${-decimals}`);
}

// The rounded total of the group: PLSFIX.ROUNDSUM, and by construction the sum of
// what allocateRounded returns for the same values.
export function roundedTotal(values: number[], decimals: number): number {
  assertInputs(values, decimals);
  return fromUnits(totalUnits(values, decimals), decimals);
}

// Indices ordered by what each value lost to the floor, largest first; equal
// remainders keep their position order, so the allocation is the same for every
// sibling cell that recomputes it. The remainders are exact integers over one
// shared power of ten, so two values that lost the same decimal amount are
// equal here whatever their size.
function bumpOrder(remainders: bigint[]): number[] {
  return remainders
    .map((remainder, index) => ({ remainder, index }))
    .sort((a, b) =>
      a.remainder === b.remainder
        ? a.index - b.index
        : a.remainder > b.remainder
          ? -1
          : 1,
    )
    .map((entry) => entry.index);
}

// Whole units the floors still owe the total: between 0 and one per value, so
// there is no such thing as bumping a value that is not there. It only leaves
// that range when the total and the parts cannot agree to the last digit -
// huge values cancelling, or more digits than Excel itself holds - and then
// the parts stay within a unit of their own values instead.
function unitsOwed(total: bigint, floored: bigint, count: number): number {
  const owed = total - floored;
  if (owed < 0n) return 0;
  return owed > BigInt(count) ? count : Number(owed);
}

// Each value rounded to `decimals`, with the whole set adding up to
// roundedTotal(values, decimals) exactly. Deterministic: same inputs, same
// output, in any cell that asks.
export function allocateRounded(values: number[], decimals: number): number[] {
  assertInputs(values, decimals);
  if (values.length === 0) return [];

  const parts = values.map((value) => unitsFloored(held(value), decimals));
  const places = parts.reduce((most, part) => Math.max(most, part.places), 0);
  const floored = parts.reduce((sum, part) => sum + part.floor, 0n);
  const owed = unitsOwed(totalUnits(values, decimals), floored, values.length);

  const bumped = new Set(
    bumpOrder(
      parts.map((part) => part.dropped * tenPower(places - part.places)),
    ).slice(0, owed),
  );
  return parts.map((part, index) =>
    fromUnits(part.floor + (bumped.has(index) ? 1n : 0n), decimals),
  );
}
