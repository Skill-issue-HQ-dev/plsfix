// The three-statement mini block against the accounting it names: an income
// statement, a balance sheet that opens on typed balances and a cash flow, with
// interest on the opening net debt and cash as the plug. Computed by
// test/formula-eval.ts; the direct method on four businesses is in the .cases file.
import { describe, expect, it } from "vitest";
import { cellAt, errorsIn } from "../../test/block-grid";
import { evaluateTemplate } from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { entryById, TEMPLATES } from "./index";
import { filterTemplates } from "./search";
import { threeStatement } from "./three-statement";

// Column index 2, row index 4: the corner the pinned formulas resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: the opening balances in B, years one to three
// in C to E.
const OPENING = "B";
const YEARS = ["C", "D", "E"] as const;
const ROW = {
  header: 2,
  revenue: 3,
  margin: 4,
  ebitda: 5,
  da: 6,
  ebit: 7,
  rate: 8,
  interest: 9,
  ebt: 10,
  taxRate: 11,
  tax: 12,
  netIncome: 13,
  ppe: 15,
  wc: 16,
  cash: 17,
  debt: 18,
  equity: 19,
  netDebt: 20,
  check: 21,
  plusDa: 23,
  capex: 24,
  wcIncrease: 25,
  fcf: 26,
} as const;

const addr = (column: string, row: number): string => `${column}${String(row)}`;

const block = threeStatement();

describe("the three-statement mini block", () => {
  it("is a 26 by 5 block named Three-statement mini, titled by its name", () => {
    expect(block.id).toBe("three-statement");
    expect(block.name).toBe("Three-statement mini");
    expect([block.rows, block.cols]).toEqual([26, 5]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Three-statement mini",
    });
  });

  it("heads an Opening column in text, then numbers the three years", () => {
    expect(cellAt(block, "A2")).toMatchObject({ kind: "header", v: "Year" });
    expect(cellAt(block, "B2")).toMatchObject({ kind: "header", v: "Opening" });
    YEARS.forEach((column, index) => {
      expect(cellAt(block, addr(column, ROW.header)), column).toMatchObject({
        kind: "header",
        v: index + 1,
      });
    });
  });

  it("names the unit of every row, and says what the check checks", () => {
    const name = (row: number): string =>
      String(cellAt(block, addr("A", row))?.v);
    const rates = [ROW.margin, ROW.rate, ROW.taxRate];
    const unitless: number[] = [...rates, ROW.check, ROW.header];
    for (const row of Object.values(ROW)) {
      if (!unitless.includes(row))
        expect(name(row), `row ${String(row)}`).toMatch(/EUR m/);
    }
    for (const row of rates) expect(name(row)).toMatch(/%/);
    expect(name(ROW.rate)).toMatch(/opening net debt/);
    expect(name(ROW.check)).toMatch(/^Check:/);
    expect(name(ROW.check)).toMatch(/balance|assets/i);
  });

  it("types the assumptions and the opening balances in, and builds every other line with a formula", () => {
    const shapes: [number[], { kind: string; format: string }][] = [
      [
        [ROW.revenue, ROW.da, ROW.capex, ROW.wcIncrease],
        { kind: "input", format: "decimal" },
      ],
      [
        [ROW.margin, ROW.rate, ROW.taxRate],
        { kind: "input", format: "percent" },
      ],
      [
        [ROW.ebitda, ROW.ebit, ROW.interest, ROW.ebt, ROW.tax, ROW.plusDa],
        { kind: "formula", format: "decimal" },
      ],
      [[ROW.netIncome, ROW.fcf], { kind: "result", format: "decimal" }],
    ];
    for (const [rows, shape] of shapes) {
      for (const row of rows) {
        for (const column of YEARS) {
          const address = addr(column, row);
          expect(cellAt(block, address), address).toMatchObject(shape);
        }
        // The Opening column holds nothing on these lines; a result line keeps
        // its band across the empty cell.
        const empty = shape.kind === "result" ? "result" : "label";
        expect(cellAt(block, addr(OPENING, row)), `B${row}`).toEqual({
          kind: empty,
        });
      }
    }
    // A balance is typed in the Opening column and rolled in each year after it.
    for (const row of [ROW.ppe, ROW.wc, ROW.cash, ROW.debt, ROW.equity]) {
      expect(cellAt(block, addr(OPENING, row)), `B${row}`).toMatchObject({
        kind: "input",
        format: "decimal",
      });
      for (const column of YEARS) {
        expect(
          cellAt(block, addr(column, row)),
          addr(column, row),
        ).toMatchObject({ kind: "formula", format: "decimal" });
      }
    }
    // The memo and the check are formulas in all four columns.
    for (const row of [ROW.netDebt, ROW.check]) {
      for (const column of [OPENING, ...YEARS]) {
        expect(
          cellAt(block, addr(column, row)),
          addr(column, row),
        ).toMatchObject({ kind: "formula", format: "decimal" });
      }
    }
  });

  it("works the demo through its three statements", () => {
    const sheet = evaluateTemplate(block);
    // by hand, year one: EBITDA 120 x 20 % = 24; EBIT 24 - 6 = 18; interest 7.5 %
    // of the opening net debt 50 - 10 = 40 is 3; EBT 15; tax 3; net income 12;
    // FCF 12 + 6 - 8 - 2 = 8; cash 18; net debt 50 - 18 = 32; equity 52; PP&E 62.
    // Year two: interest 7.5 % of 32 = 2.4; EBITDA 27.72; EBIT 21.22; EBT 18.82;
    // tax 3.764; net income 15.056; FCF 10.056; cash 28.056; net debt 21.944;
    // equity 67.056. Year three: interest 1.6458; EBITDA 31.68; EBIT 24.68; EBT
    // 23.0342; tax 4.60684; net income 18.42736; FCF 13.42736; cash 41.48336; net
    // debt 8.51664; equity 85.48336.
    const want: Record<number, number[]> = {
      [ROW.ebitda]: [24, 27.72, 31.68],
      [ROW.ebit]: [18, 21.22, 24.68],
      [ROW.interest]: [3, 2.4, 1.6458],
      [ROW.ebt]: [15, 18.82, 23.0342],
      [ROW.tax]: [3, 3.764, 4.60684],
      [ROW.netIncome]: [12, 15.056, 18.42736],
      [ROW.plusDa]: [6, 6.5, 7],
      [ROW.fcf]: [8, 10.056, 13.42736],
      [ROW.ppe]: [62, 64.5, 67],
      [ROW.wc]: [22, 24.5, 27],
      [ROW.cash]: [18, 28.056, 41.48336],
      [ROW.debt]: [50, 50, 50],
      [ROW.equity]: [52, 67.056, 85.48336],
      [ROW.netDebt]: [32, 21.944, 8.51664],
      [ROW.check]: [0, 0, 0],
    };
    for (const [row, figures] of Object.entries(want)) {
      YEARS.forEach((column, index) => {
        expect(
          sheet.value(addr(column, Number(row))),
          `${column}${row}`,
        ).toBeCloseTo(figures[index] as number, 9);
      });
    }
    // The opening column: net debt 50 - 10 = 40, and 60 + 20 + 10 - 50 - 40 = 0.
    expect(sheet.value(addr(OPENING, ROW.netDebt))).toBe(40);
    expect(sheet.value(addr(OPENING, ROW.check))).toBe(0);
  });

  it("holds debt flat: each year takes the debt of the year before, so a typed-over year carries on", () => {
    const sheet = evaluateTemplate(block, { C18: 55 });
    expect(sheet.value("D18")).toBe(55);
    expect(sheet.value("E18")).toBe(55);
    expect(evaluateTemplate(block).value("E18")).toBe(50);
  });

  it("computes with no error and no circular reference, with or without a loss", () => {
    expect(errorsIn(block)).toEqual([]);
    expect(errorsIn(block, { C3: 10, D3: 0, E3: null })).toEqual([]);
  });

  it("writes interest, tax, the rolls, the cash flow and the check where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(8, 2)).toBe("=E12*D24");
    expect(formula(8, 3)).toBe("=F12*E24");
    expect(formula(11, 2)).toBe("=MAX(E14,0)*E15");
    expect(formula(14, 2)).toBe("=D19+E28-E10");
    expect(formula(15, 2)).toBe("=D20+E29");
    expect(formula(16, 2)).toBe("=D21+E30");
    expect(formula(17, 2)).toBe("=D22");
    expect(formula(18, 2)).toBe("=D23+E17");
    expect(formula(19, 1)).toBe("=D22-D21");
    expect(formula(25, 2)).toBe("=E17+E27-E28-E29");
    expect(formula(20, 2)).toBe(
      "=ROUND(ABS(E19+E20+E21-E22-E23)+ABS(E24-(E22-E21)),6)",
    );
  });
});

describe("the three-statement mini in the picker", () => {
  it("is listed under Operating, carrying this very block", () => {
    const entry = entryById("template-three-statement");
    expect(entry?.group).toBe("Operating");
    expect(entry?.block).toEqual(block);
  });

  it("is found by the words a modeller types for it", () => {
    for (const query of [
      "three statement",
      "income statement",
      "balance sheet",
      "cash flow",
    ]) {
      const ids = filterTemplates(TEMPLATES, query).map((entry) => entry.id);
      expect(ids, query).toContain("template-three-statement");
    }
  });
});
