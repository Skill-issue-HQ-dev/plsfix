// The function library of the test-only formula evaluator: the lazy IF, IFERROR
// and CHOOSE, ROUND, N, ISTEXT and PERCENTILE.INC, plus ARITY, the one table of
// every function the evaluator knows and the argument counts Excel accepts.
// Owns Fault and the coercions; formula-eval.ts owns the grid and the aggregates.

import type { Node } from "./formula-syntax";

export type Value = number | string | boolean | null;

/** An error raised mid-formula; a cell turns it into an ExcelError value. */
export class Fault extends Error {
  readonly code: string;
  constructor(code: string) {
    super(code);
    this.code = code;
  }
}

export function toNumber(value: Value): number {
  if (value === null) return 0;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value.trim() !== "" && !Number.isNaN(Number(value))) return Number(value);
  throw new Fault("#VALUE!");
}

export function toBoolean(value: Value): boolean {
  if (typeof value === "boolean") return value;
  if (value === null) return false;
  if (typeof value === "number") return value !== 0;
  throw new Fault("#VALUE!");
}

/** What a function needs from the evaluator that calls it. */
export interface Reader {
  /** One argument's value; an error it raises travels as a Fault. */
  value(arg: Node): Value;
  /** The numbers an aggregate takes from these arguments, by reference rules. */
  numbers(name: string, args: Node[]): number[];
}

// Excel refuses a formula whose call has too few or too many arguments, so
// the evaluator does too, by name, instead of reading around the gap.
const MANY = Number.POSITIVE_INFINITY;
export const ARITY: Record<string, [number, number]> = {
  IF: [2, 3],
  IFERROR: [2, 2],
  CHOOSE: [2, 255],
  ROUND: [2, 2],
  AND: [1, MANY],
  OR: [1, MANY],
  ABS: [1, 1],
  N: [1, 1],
  ISTEXT: [1, 1],
  "PERCENTILE.INC": [2, 2],
  SUM: [1, MANY],
  AVERAGE: [1, MANY],
  MEDIAN: [1, MANY],
  MAX: [1, MANY],
  MIN: [1, MANY],
  COUNT: [1, MANY],
  IRR: [1, 2],
};

export function checkArity(name: string, count: number): void {
  const limits = ARITY[name];
  if (limits === undefined) {
    throw new Error(`formula-eval does not know the function ${name}`);
  }
  if (count < limits[0] || count > limits[1]) {
    throw new Error(
      `${name} takes ${limits[0]} to ${limits[1]} arguments, got ${count}`,
    );
  }
}

type Library = (args: Node[], reader: Reader) => Value;

function ifError(args: Node[], reader: Reader): Value {
  try {
    return reader.value(args[0] as Node);
  } catch (error) {
    if (!(error instanceof Fault)) throw error;
    return reader.value(args[1] as Node);
  }
}

function condition(args: Node[], reader: Reader): Value {
  const taken = toBoolean(reader.value(args[0] as Node)) ? args[1] : args[2];
  return taken === undefined ? false : reader.value(taken);
}

// CHOOSE counts from one, drops the fraction of its index and reads only the
// value it picks, so an error in another value stays hidden.
function choose(args: Node[], reader: Reader): Value {
  const index = Math.trunc(toNumber(reader.value(args[0] as Node)));
  const picked = args[index];
  if (index < 1 || picked === undefined) throw new Fault("#VALUE!");
  return reader.value(picked);
}

// ROUND works on what a cell displays: the scaled value is cut to 15
// significant digits first (1.005 is 1.005 to Excel, not 1.00499999999999989),
// then rounded half away from zero. A result that is zero is never minus zero.
// Exported on its own so a suite can hold another rounding to Excel's answer.
export function excelRound(number: number, digits: number): number {
  const shift = (x: number): number =>
    digits >= 0 ? x * 10 ** digits : x / 10 ** -digits;
  const unshift = (x: number): number =>
    digits >= 0 ? x / 10 ** digits : x * 10 ** -digits;
  const shifted = Number(shift(Math.abs(number)).toPrecision(15));
  const rounded = Math.sign(number) * unshift(Math.round(shifted));
  return rounded === 0 ? 0 : rounded;
}

function round(args: Node[], reader: Reader): Value {
  const number = toNumber(reader.value(args[0] as Node));
  return excelRound(
    number,
    Math.trunc(toNumber(reader.value(args[1] as Node))),
  );
}

// N reads a number as itself, a logical as 1 or 0 and all else as 0; an error
// stays an error. ISTEXT never raises one: an error is simply not text.
function numberOf(args: Node[], reader: Reader): Value {
  const value = reader.value(args[0] as Node);
  if (typeof value === "number") return value;
  return typeof value === "boolean" && value ? 1 : 0;
}

function isText(args: Node[], reader: Reader): Value {
  try {
    return typeof reader.value(args[0] as Node) === "string";
  } catch (error) {
    if (error instanceof Fault) return false;
    throw error;
  }
}

// PERCENTILE.INC: the k-th share of the way along the sorted numbers, the two
// nearest values interpolated; k outside 0..1 or nothing to read is #NUM!.
function percentile(args: Node[], reader: Reader): Value {
  const k = toNumber(reader.value(args[1] as Node));
  const numbers = reader.numbers("PERCENTILE.INC", [args[0] as Node]);
  if (numbers.length === 0 || k < 0 || k > 1) throw new Fault("#NUM!");
  const sorted = [...numbers].sort((a, b) => a - b);
  const position = k * (sorted.length - 1);
  const below = Math.floor(position);
  const lower = sorted[below] ?? 0;
  const upper = sorted[Math.ceil(position)] ?? lower;
  return lower + (position - below) * (upper - lower);
}

const LIBRARY: Record<string, Library> = {
  IF: condition,
  IFERROR: ifError,
  CHOOSE: choose,
  ROUND: round,
  N: numberOf,
  ISTEXT: isText,
  "PERCENTILE.INC": percentile,
};

/**
 * Answers a call to one of this library's functions after checking its argument
 * count; undefined when the name is not one of them (the evaluator's own AND, OR,
 * ABS and aggregates), which no function here ever answers.
 */
export function callLibrary(
  name: string,
  args: Node[],
  reader: Reader,
): Value | undefined {
  checkArity(name, args.length);
  return LIBRARY[name]?.(args, reader);
}
