// The football-field data block against the chart that reads it: laid out
// exactly as the Football field button expects its selection, so a modeller
// inserts the block, presses the button with the block still selected and gets
// the chart. The proof runs the chart's own selection parser (readTriples) and
// its range maths (footballField) over the block's grid as Excel would hand it
// over, and checks the midpoints the block computes itself.
import { describe, expect, it } from "vitest";
import { evaluateTemplate, ExcelError } from "../../test/formula-eval";
import { cellAddress } from "../find";
import { footballField as chartRanges } from "../chartmath";
import { readTriples, type TripleRules } from "../excel/chart-blocks";
import { type CellValue } from "../model";
import { resolveFormula } from "../template-cells";
import { footballField } from "./football-field";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// What the Football field button demands of a selection, restated: at least
// two rows and a label, a low and a high on each (src/excel/football.ts).
const BUTTON_RULES: TripleRules = {
  minRows: 2,
  tooFew: "need at least two rows",
  notNumbers: "the low and high columns must hold numbers",
};

const METHODS = [
  "Trading comps",
  "Precedents",
  "DCF",
  "52-week range",
  "Analyst targets",
];

const block = footballField();

// The block as a selection's `values`: a literal as typed, a formula as it
// computes, an empty cell as the empty string Excel answers for it.
function selectionValues(
  overrides: Record<string, number> = {},
): CellValue[][] {
  const sheet = evaluateTemplate(block, overrides);
  return block.cells.map((row, r) =>
    row.map((_cell, c) => {
      const value = sheet.value(cellAddress(r, c));
      if (value === null) return "";
      if (value instanceof ExcelError) return value.code;
      return value;
    }),
  );
}

describe("the football field data block", () => {
  it("is a 6 by 4 block named Football field data", () => {
    expect(block.id).toBe("football-field");
    expect(block.name).toBe("Football field data");
    expect([block.rows, block.cols]).toEqual([6, 4]);
  });

  it("opens on its header row, with no title above it for the parser to trip on", () => {
    expect(block.cells[0]?.map((cell) => cell.v)).toEqual([
      "Method (EUR m)",
      "Low",
      "High",
      "Midpoint",
    ]);
    expect(block.cells[0]?.every((cell) => cell.kind === "header")).toBe(true);
    expect(
      block.cells.some((row) => row.some((cell) => cell.kind === "title")),
    ).toBe(false);
  });

  it("lists the five methods a football field shows, in that order", () => {
    expect(block.cells.slice(1).map((row) => row[0]?.v)).toEqual(METHODS);
  });

  it("types each method's low and high in as inputs and computes the midpoint", () => {
    for (const row of block.cells.slice(1)) {
      expect(row[1]).toMatchObject({ kind: "input", format: "decimal" });
      expect(row[2]).toMatchObject({ kind: "input", format: "decimal" });
      expect(row[3]).toMatchObject({ kind: "formula", format: "decimal" });
      expect(typeof row[1]?.v).toBe("number");
      expect(typeof row[2]?.v).toBe("number");
    }
  });

  it("is read whole by the chart's selection parser: header detected, five rows", () => {
    const rows = readTriples(selectionValues(), BUTTON_RULES);
    expect(rows.map((row) => row.label)).toEqual(METHODS);
    for (const row of rows) {
      expect(row.low, row.label).toBeGreaterThan(0);
      expect(row.high, row.label).toBeGreaterThan(row.low);
    }
  });

  it("hands the chart ranges with no row swapped and a band equal to high less low", () => {
    const rows = readTriples(selectionValues(), BUTTON_RULES);
    const ranges = chartRanges(rows);
    expect(ranges.swaps).toBe(0);
    expect(ranges.labels).toEqual(METHODS);
    ranges.range.forEach((band, index) => {
      expect(band).toBeCloseTo(
        (ranges.high[index] as number) - (ranges.low[index] as number),
        12,
      );
    });
  });

  it("would not parse with a title row above the header, which is why it has none", () => {
    const titled: CellValue[][] = [
      ["Football field data", "", "", ""],
      ...selectionValues(),
    ];
    expect(() => readTriples(titled, BUTTON_RULES)).toThrow(
      BUTTON_RULES.notNumbers,
    );
  });

  it("computes each midpoint as the average of its low and high", () => {
    const values = selectionValues();
    for (const row of values.slice(1)) {
      expect(row[3]).toBeCloseTo(
        ((row[1] as number) + (row[2] as number)) / 2,
        12,
      );
    }
  });

  it("follows a typed low and high to a new midpoint", () => {
    const values = selectionValues({ B2: 100, C2: 200 });
    expect(values[1]?.[3]).toBe(150);
  });

  it("demonstrates a football field: ranges that share a zone every method agrees on", () => {
    const rows = readTriples(selectionValues(), BUTTON_RULES);
    expect(Math.max(...rows.map((row) => row.low))).toBeLessThan(
      Math.min(...rows.map((row) => row.high)),
    );
  });

  it("writes the midpoint formulas against their own row", () => {
    const formula = (row: number): string | undefined => {
      const cell = block.cells[row]?.[3];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(1)).toBe("=AVERAGE(D6:E6)");
    expect(formula(5)).toBe("=AVERAGE(D10:E10)");
  });
});
