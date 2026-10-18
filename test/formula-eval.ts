// A small Excel formula evaluator for the template suites: references, ranges,
// Excel's operator levels, comparisons, the aggregates, IRR, AND, OR and ABS;
// the other functions sit in formula-functions.ts. Test support only: the fake
// host has no formula engine, and a function it does not know throws by name.

import { cellAddress } from "../src/find";
import { resolveFormula, type Template } from "../src/template-cells";
import { irr } from "./formula-irr";
import { type Node, parseFormula } from "./formula-syntax";
import {
  callLibrary,
  Fault,
  type Reader,
  toBoolean,
  toNumber,
  type Value,
} from "./formula-functions";

export type { Value } from "./formula-functions";

/** An Excel error value (#DIV/0!, #VALUE!, #NUM!) a cell can hold. */
export class ExcelError {
  readonly code: string;
  constructor(code: string) {
    this.code = code;
  }
}

export type Computed = Value | ExcelError;

// ---------------------------------------------------------------------------
// Values
// ---------------------------------------------------------------------------

// Numbers sort before text; a blank takes the zero or the empty text of what it
// sits against; text compares without case.
function compare(left: Value, right: Value): number {
  const blankAs = (other: Value): Value => (typeof other === "string" ? "" : 0);
  const a = left === null ? blankAs(right) : left;
  const b = right === null ? blankAs(left) : right;
  if (typeof a === "string" && typeof b === "string") {
    return a.toLowerCase().localeCompare(b.toLowerCase());
  }
  if (typeof a === "string") return 1;
  if (typeof b === "string") return -1;
  return toNumber(a) - toNumber(b);
}

/** The row and column (both from 0) of an A1 reference, $ and case ignored. */
export function parseRef(ref: string): { row: number; column: number } {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)$/.exec(ref.toUpperCase());
  if (!match) throw new Error(`not a cell reference: ${ref}`);
  let column = 0;
  for (const letter of match[1] ?? "")
    column = column * 26 + letter.charCodeAt(0) - 64;
  return { row: Number(match[2]) - 1, column: column - 1 };
}

// ---------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------

type CellReader = (address: string) => Computed;

function cellValue(read: CellReader, address: string): Value {
  const cell = read(address);
  if (cell instanceof ExcelError) throw new Fault(cell.code);
  return cell;
}

/** The A1 address of every cell of a range, row by row. */
export function rangeAddresses(from: string, to: string): string[] {
  const a = parseRef(from);
  const b = parseRef(to);
  const addresses: string[] = [];
  for (
    let row = Math.min(a.row, b.row);
    row <= Math.max(a.row, b.row);
    row += 1
  ) {
    for (
      let column = Math.min(a.column, b.column);
      column <= Math.max(a.column, b.column);
      column += 1
    ) {
      addresses.push(cellAddress(row, column));
    }
  }
  return addresses;
}

/** Every cell of a range, row by row, an error kept as the value it is. */
function rangeCells(read: CellReader, from: string, to: string): Computed[] {
  return rangeAddresses(from, to).map((address) => read(address));
}

// What an aggregate takes from a cell it is handed by reference, one cell or a
// range alike: a number counts; text, blanks and logicals are skipped; an error
// travels, except that COUNT steps over it.
function numberFromCell(cell: Computed, name: string): number[] {
  if (cell instanceof ExcelError) {
    if (name === "COUNT") return [];
    throw new Fault(cell.code);
  }
  return typeof cell === "number" ? [cell] : [];
}

// A value typed into the call is coerced the way arithmetic coerces it ("3" is
// 3, a logical is 1 or 0, any other text is #VALUE!); COUNT leaves out what it
// cannot read instead.
function typedNumber(arg: Node, read: CellReader, name: string): number[] {
  try {
    return [toNumber(evaluate(arg, read))];
  } catch (error) {
    if (name === "COUNT" && error instanceof Fault) return [];
    throw error;
  }
}

function numbersIn(read: CellReader, name: string, args: Node[]): number[] {
  return args.flatMap((arg) => {
    if (arg.type === "range") {
      const cells = rangeCells(read, arg.from, arg.to);
      return cells.flatMap((cell) => numberFromCell(cell, name));
    }
    if (arg.type === "ref") return numberFromCell(read(arg.address), name);
    return typedNumber(arg, read, name);
  });
}

function median(numbers: number[]): number {
  if (numbers.length === 0) throw new Fault("#NUM!");
  const sorted = [...numbers].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 1
    ? (sorted[middle] ?? 0)
    : ((sorted[middle - 1] ?? 0) + (sorted[middle] ?? 0)) / 2;
}

function average(numbers: number[]): number {
  if (numbers.length === 0) throw new Fault("#DIV/0!");
  return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
}

function aggregate(name: string, numbers: number[]): number {
  switch (name) {
    case "SUM":
      return numbers.reduce((sum, value) => sum + value, 0);
    case "AVERAGE":
      return average(numbers);
    case "MEDIAN":
      return median(numbers);
    case "MAX":
      return numbers.length === 0 ? 0 : Math.max(...numbers);
    case "MIN":
      return numbers.length === 0 ? 0 : Math.min(...numbers);
    case "COUNT":
      return numbers.length;
    default:
      // A name that reaches here without a case would otherwise be counted.
      throw new Error(`formula-eval has no rule for ${name}`);
  }
}

// IRR reads the numbers of its first argument (text and blanks are skipped, so a blank
// year squeezes the periods, as in Excel) and an optional guess; no root is #NUM!.
function irrOf(args: Node[], read: CellReader): number {
  const guess =
    args[1] === undefined ? undefined : toNumber(evaluate(args[1], read));
  const rate = irr(numbersIn(read, "IRR", args.slice(0, 1)), guess);
  if (rate === null) throw new Fault("#NUM!");
  return rate;
}

// AND and OR read every argument, as Excel does, so an error in one still shows.
function logical(name: string, args: Node[], read: CellReader): boolean {
  const truths = args.map((arg) => toBoolean(evaluate(arg, read)));
  return name === "AND" ? truths.every(Boolean) : truths.some(Boolean);
}

// The library answers its own functions (arity checked first); IRR, AND, OR, ABS
// and the aggregates are the evaluator's.
function call(name: string, args: Node[], read: CellReader): Value {
  const reader: Reader = {
    value: (arg) => evaluate(arg, read),
    numbers: (fn, list) => numbersIn(read, fn, list),
  };
  const answer = callLibrary(name, args, reader);
  if (answer !== undefined) return answer;
  if (name === "AND" || name === "OR") return logical(name, args, read);
  if (name === "IRR") return irrOf(args, read);
  if (name === "ABS")
    return Math.abs(toNumber(evaluate(args[0] as Node, read)));
  return aggregate(name, numbersIn(read, name, args));
}

// Excel's ^: 0 to a negative power is #DIV/0!; 0^0, a root of a negative and an overflow are #NUM!.
function power(base: number, exponent: number): number {
  if (base === 0 && exponent < 0) throw new Fault("#DIV/0!");
  if (base === 0 && exponent === 0) throw new Fault("#NUM!");
  const result = base ** exponent;
  if (!Number.isFinite(result)) throw new Fault("#NUM!");
  return result;
}

function arithmetic(op: string, left: number, right: number): number {
  switch (op) {
    case "+":
      return left + right;
    case "-":
      return left - right;
    case "*":
      return left * right;
    case "^":
      return power(left, right);
    default:
      if (right === 0) throw new Fault("#DIV/0!");
      return left / right;
  }
}

function comparison(op: string, left: Value, right: Value): boolean {
  const order = compare(left, right);
  const answers: Record<string, boolean> = {
    "=": order === 0,
    "<>": order !== 0,
    "<": order < 0,
    ">": order > 0,
    "<=": order <= 0,
    ">=": order >= 0,
  };
  return answers[op] ?? false;
}

function evaluate(node: Node, read: CellReader): Value {
  switch (node.type) {
    case "num":
    case "str":
      return node.value;
    case "ref":
      return cellValue(read, node.address);
    case "range":
      throw new Fault("#VALUE!");
    case "neg":
      return -toNumber(evaluate(node.operand, read));
    case "call":
      return call(node.name, node.args, read);
    case "bin": {
      const left = evaluate(node.left, read);
      const right = evaluate(node.right, read);
      return ["=", "<>", "<", ">", "<=", ">="].includes(node.op)
        ? comparison(node.op, left, right)
        : arithmetic(node.op, toNumber(left), toNumber(right));
    }
  }
}

// ---------------------------------------------------------------------------
// A block as a sheet
// ---------------------------------------------------------------------------

export interface Evaluated {
  /** What the cell at an A1 address shows, the block's corner being A1. */
  value(address: string): Computed;
}

/**
 * Computes every cell of a block placed at A1. `overrides` replace a cell's own
 * content with a typed literal, the way a modeller overtypes an input; a
 * literal may be an ExcelError, the way a cell can hold one.
 */
export function evaluateTemplate(
  template: Template,
  overrides: Record<string, Computed> = {},
): Evaluated {
  const source = new Map<string, { formula?: string; literal?: Computed }>();
  const corner = { row: 0, column: 0 };
  template.cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      const address = cellAddress(r, c);
      if (cell.f !== undefined) {
        source.set(address, { formula: resolveFormula(cell.f, corner) });
      } else if (cell.v !== undefined) {
        source.set(address, { literal: cell.v });
      }
    }),
  );
  for (const [address, literal] of Object.entries(overrides)) {
    source.set(address, { literal });
  }

  const cache = new Map<string, Computed>();
  const working = new Set<string>();
  function value(address: string): Computed {
    const key = address.replace(/\$/g, "").toUpperCase();
    if (cache.has(key)) return cache.get(key) ?? null;
    if (working.has(key)) throw new Error(`circular reference at ${key}`);
    working.add(key);
    const entry = source.get(key);
    let result: Computed = entry?.literal ?? null;
    if (entry?.formula !== undefined) {
      try {
        const tree = parseFormula(entry.formula.slice(1));
        result = evaluate(tree, value) ?? 0; // Excel: a formula is never blank
      } catch (error) {
        if (!(error instanceof Fault)) throw error;
        result = new ExcelError(error.code);
      }
    }
    working.delete(key);
    cache.set(key, result);
    return result;
  }
  return { value };
}
