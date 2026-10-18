// Test support for the template blocks' own suites: a block cell by its A1
// address (the block's corner being A1), every error value a block computes and
// findLoop, the static walk over the formulas' references. Reading every cell
// proves no loop on the path the values take; only findLoop proves there is none.

import { cellAddress } from "../src/find";
import {
  resolveFormula,
  type Template,
  type TemplateCell,
} from "../src/template-cells";
import {
  type Computed,
  evaluateTemplate,
  ExcelError,
  parseRef,
  rangeAddresses,
} from "./formula-eval";
import { type Node, parseFormula } from "./formula-syntax";

const CORNER = { row: 0, column: 0 };

/** The block's cell at an A1 address, undefined past the block's edge. */
export function cellAt(
  template: Template,
  address: string,
): TemplateCell | undefined {
  const { row, column } = parseRef(address);
  return template.cells[row]?.[column];
}

/** Every cell whose computed value is an Excel error, as "B5 #DIV/0!". */
export function errorsIn(
  template: Template,
  overrides: Record<string, Computed> = {},
): string[] {
  const sheet = evaluateTemplate(template, overrides);
  return template.cells.flatMap((row, r) =>
    row.flatMap((_cell, c) => {
      const address = cellAddress(r, c);
      const value = sheet.value(address);
      return value instanceof ExcelError ? [`${address} ${value.code}`] : [];
    }),
  );
}

// Every cell a formula names anywhere in its text, a range as the cells it
// covers: lazy functions do not hide a reference here, as they do in a value.
function referencesIn(node: Node): string[] {
  switch (node.type) {
    case "ref": {
      const { row, column } = parseRef(node.address);
      return [cellAddress(row, column)];
    }
    case "range":
      return rangeAddresses(node.from, node.to);
    case "neg":
      return referencesIn(node.operand);
    case "bin":
      return [...referencesIn(node.left), ...referencesIn(node.right)];
    case "call":
      return node.args.flatMap(referencesIn);
    default:
      return [];
  }
}

// Each formula cell's address with the cells its formula names.
function precedents(template: Template): Map<string, string[]> {
  const graph = new Map<string, string[]>();
  template.cells.forEach((row, r) =>
    row.forEach((cell, c) => {
      if (cell.f === undefined) return;
      const source = resolveFormula(cell.f, CORNER).slice(1);
      graph.set(cellAddress(r, c), referencesIn(parseFormula(source)));
    }),
  );
  return graph;
}

/**
 * The cells round one dependency loop of a block, the first repeated at the end
 * ("A1", "A2", "A1"), or [] for none. A depth-first walk over every cell a
 * formula names, taken or not: a cell met again on the walk's own path is a back
 * edge, a loop; a typed value or a label ends the walk.
 */
export function findLoop(template: Template): string[] {
  const graph = precedents(template);
  const finished = new Set<string>();
  const path: string[] = [];

  function visit(address: string): string[] {
    if (finished.has(address)) return [];
    const seen = path.indexOf(address);
    if (seen >= 0) return [...path.slice(seen), address];
    path.push(address);
    for (const next of graph.get(address) ?? []) {
      const loop = visit(next);
      if (loop.length > 0) return loop;
    }
    path.pop();
    finished.add(address);
    return [];
  }

  for (const address of graph.keys()) {
    const loop = visit(address);
    if (loop.length > 0) return loop;
  }
  return [];
}
