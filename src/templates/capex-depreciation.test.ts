// The capex-and-depreciation block against the accounting it names: a PP&E
// roll-forward in gross and net, straight-line on the gross balance, half a year
// on the year's capex, never taking the net book value below zero. Computed by
// test/formula-eval.ts and held to a schedule worked out from sums, not the roll.
import { describe, expect, it } from "vitest";
import { cellAt, errorsIn } from "../../test/block-grid";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { capexDepreciation } from "./capex-depreciation";
import { entryById, TEMPLATES } from "./index";
import { filterTemplates } from "./search";

// Column index 2, row index 4: the corner the pinned formulas resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: the useful life in B2, years one to five in B to F.
const YEARS = ["B", "C", "D", "E", "F"] as const;
const LIFE = "B2";
const OPENING_GROSS = "B9";
const OPENING_NET = "B12";
const ROW = {
  header: 4,
  revenue: 5,
  capex: 6,
  capexShare: 7,
  grossOpening: 9,
  grossClosing: 10,
  netOpening: 12,
  plusCapex: 13,
  depreciation: 14,
  netClosing: 15,
} as const;

const addr = (column: string, row: number): string => `${column}${String(row)}`;

interface Case {
  life: number;
  gross: number;
  net: number;
  revenue: number[];
  capex: number[];
}

const DEMO: Case = {
  life: 10,
  gross: 200,
  net: 200,
  revenue: [300, 320, 340, 360, 380],
  capex: [30, 33, 36, 38, 40],
};

// The cases the schedule is held to: the demo and others typed over it, none of
// them short-lived enough for the net balance to run out.
const CASES: [string, Case][] = [
  ["the demo case", DEMO],
  [
    "a long life and heavy early capex",
    {
      life: 25,
      gross: 500,
      net: 500,
      revenue: [900, 950, 1000, 1050, 1100],
      capex: [120, 80, 60, 40, 20],
    },
  ],
  [
    "a net balance below gross, depreciation already taken",
    {
      life: 10,
      gross: 350,
      net: 200,
      revenue: [300, 320, 340, 360, 380],
      capex: [30, 0, 36, 0, 40],
    },
  ],
  [
    "no opening balance",
    {
      life: 8,
      gross: 0,
      net: 0,
      revenue: [200, 220, 240, 260, 280],
      capex: [100, 100, 100, 100, 100],
    },
  ],
];

// A case as the cells a modeller overtypes.
function typed(c: Case): Record<string, Value> {
  const cells: Record<string, Value> = {
    [LIFE]: c.life,
    [OPENING_GROSS]: c.gross,
    [OPENING_NET]: c.net,
  };
  YEARS.forEach((column, index) => {
    cells[addr(column, ROW.revenue)] = c.revenue[index] ?? null;
    cells[addr(column, ROW.capex)] = c.capex[index] ?? null;
  });
  return cells;
}

// Straight-line on gross, from sums rather than from the roll: the opening gross
// and every earlier year's capex are written off 1 / life a year, the year's own
// capex half of that, and a charge never takes the net balance below zero.
function scheduleOf(c: Case): {
  depreciation: number[];
  gross: number[];
  net: number[];
} {
  const depreciation: number[] = [];
  const gross: number[] = [];
  const net: number[] = [];
  c.capex.forEach((capex, index) => {
    const earlier = c.capex.slice(0, index).reduce((sum, x) => sum + x, 0);
    const open = index === 0 ? c.net : (net[index - 1] ?? 0);
    const level = (c.gross + earlier) / c.life + capex / c.life / 2;
    const charge = Math.min(level, open + capex);
    depreciation.push(charge);
    gross.push(c.gross + earlier + capex);
    net.push(open + capex - charge);
  });
  return { depreciation, gross, net };
}

const block = capexDepreciation();

function expectRow(
  sheet: ReturnType<typeof evaluateTemplate>,
  row: number,
  expected: number[],
  what: string,
): void {
  YEARS.forEach((column, index) => {
    expect(sheet.value(addr(column, row)), `${what} ${column}`).toBeCloseTo(
      expected[index] as number,
      9,
    );
  });
}

describe("the capex and depreciation block", () => {
  it("is a 15 by 6 block named Capex and depreciation, titled by its name", () => {
    expect(block.id).toBe("capex-depreciation");
    expect(block.name).toBe("Capex and depreciation");
    expect([block.rows, block.cols]).toEqual([15, 6]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Capex and depreciation",
    });
  });

  it("numbers the five years in its header row", () => {
    const yearLabel = { kind: "header", v: "Year" };
    expect(cellAt(block, addr("A", ROW.header))).toMatchObject(yearLabel);
    YEARS.forEach((column, index) => {
      const cell = cellAt(block, addr(column, ROW.header));
      expect(cell, column).toMatchObject({ kind: "header", v: index + 1 });
    });
  });

  it("names the unit of every row, and the straight-line, half-year rule on depreciation", () => {
    const name = (row: number): string =>
      String(cellAt(block, addr("A", row))?.v);
    expect(String(cellAt(block, "A2")?.v)).toMatch(/years/);
    for (const row of [
      ROW.revenue,
      ROW.capex,
      ROW.grossOpening,
      ROW.grossClosing,
      ROW.netOpening,
      ROW.plusCapex,
      ROW.depreciation,
      ROW.netClosing,
    ]) {
      expect(name(row), `row ${String(row)}`).toMatch(/EUR m/);
    }
    expect(name(ROW.capexShare)).toMatch(/%/);
    expect(name(ROW.depreciation)).toBe(
      "Less: depreciation (EUR m, straight-line on gross, half-year on new capex)",
    );
  });

  it("types the life, both opening balances, revenue and capex in, and builds the rest", () => {
    const typedCells = [LIFE, OPENING_GROSS, OPENING_NET];
    for (const column of YEARS) {
      typedCells.push(addr(column, ROW.revenue), addr(column, ROW.capex));
    }
    for (const address of typedCells) {
      const cell = cellAt(block, address);
      expect(cell, address).toMatchObject({ kind: "input", format: "decimal" });
      expect(cell?.f, address).toBeUndefined();
    }
    const built: [number, string, string][] = [
      [ROW.capexShare, "formula", "percent"],
      [ROW.grossClosing, "formula", "decimal"],
      [ROW.plusCapex, "formula", "decimal"],
      [ROW.depreciation, "formula", "decimal"],
      [ROW.netClosing, "result", "decimal"],
    ];
    for (const [row, kind, format] of built) {
      for (const column of YEARS) {
        const cell = cellAt(block, addr(column, row));
        expect(cell, addr(column, row)).toMatchObject({ kind, format });
        expect(cell?.f, addr(column, row)).toBeDefined();
      }
    }
    for (const row of [ROW.grossOpening, ROW.netOpening]) {
      for (const column of YEARS.slice(1)) {
        expect(
          cellAt(block, addr(column, row)),
          addr(column, row),
        ).toMatchObject({ kind: "formula", format: "decimal" });
      }
    }
  });

  it("rolls gross and net forward over the demo's five years", () => {
    const sheet = evaluateTemplate(block);
    expect(errorsIn(block)).toEqual([]);
    // by hand, life 10: year one writes off 200 / 10 + 30 / 10 / 2 = 21.5, so net
    // closes at 200 + 30 - 21.5 = 208.5 and gross at 230; year two 230 / 10 + 33 /
    // 20 = 24.65 (net 216.85, gross 263); year three 26.3 + 1.8 = 28.1 (224.75,
    // 299); year four 29.9 + 1.9 = 31.8 (230.95, 337); year five 33.7 + 2 = 35.7
    // (235.25, 377).
    expectRow(sheet, ROW.depreciation, [21.5, 24.65, 28.1, 31.8, 35.7], "dep");
    expectRow(sheet, ROW.grossClosing, [230, 263, 299, 337, 377], "gross");
    expectRow(
      sheet,
      ROW.netClosing,
      [208.5, 216.85, 224.75, 230.95, 235.25],
      "net",
    );
  });

  it("opens each year on the balances the year before closed on", () => {
    const sheet = evaluateTemplate(block);
    expect(sheet.value(OPENING_GROSS)).toBe(200);
    expect(sheet.value(OPENING_NET)).toBe(200);
    YEARS.slice(1).forEach((column, index) => {
      const before = YEARS[index] as string;
      expect(sheet.value(addr(column, ROW.grossOpening)), column).toBe(
        sheet.value(addr(before, ROW.grossClosing)),
      );
      expect(sheet.value(addr(column, ROW.netOpening)), column).toBe(
        sheet.value(addr(before, ROW.netClosing)),
      );
    });
    // by hand: a year-one gross of 400 carries to 400 + 177 = 577 and writes off
    // more, so net ends lower; a year-one net of 400 adds its 200 to every close.
    const grosser = evaluateTemplate(block, { [OPENING_GROSS]: 400 });
    expect(grosser.value("F10")).toBe(577);
    expect(grosser.value("F15")).toBeLessThan(sheet.value("F15") as number);
    const richer = evaluateTemplate(block, { [OPENING_NET]: 400 });
    expect(richer.value("F15")).toBeCloseTo(435.25, 9);
  });

  it.each(CASES)("agrees with the schedule from sums on %s", (_name, c) => {
    const sheet = evaluateTemplate(block, typed(c));
    const expected = scheduleOf(c);
    expectRow(sheet, ROW.depreciation, expected.depreciation, "dep");
    expectRow(sheet, ROW.grossClosing, expected.gross, "gross");
    expectRow(sheet, ROW.netClosing, expected.net, "net");
  });

  it("writes off a tenth of a flat gross a year, the same amount every year", () => {
    // by hand, no capex, life 10: 200 / 10 = 20 a year whatever is left, so net
    // runs 180, 160, 140, 120, 100 and gross stays at 200.
    const sheet = evaluateTemplate(
      block,
      typed({ ...DEMO, capex: [0, 0, 0, 0, 0] }),
    );
    expectRow(sheet, ROW.depreciation, [20, 20, 20, 20, 20], "dep");
    expectRow(sheet, ROW.netClosing, [180, 160, 140, 120, 100], "net");
    expectRow(sheet, ROW.grossClosing, [200, 200, 200, 200, 200], "gross");
  });

  it("writes a lone year's capex off over half a year, then a tenth a year of its cost", () => {
    // by hand: 100 of capex, life 10, nothing before it: 100 / 10 / 2 = 5 in the
    // year it is spent, leaving 95; every later year 100 / 10 = 10, so 85 next.
    const lone = typed({
      ...DEMO,
      gross: 0,
      net: 0,
      capex: [100, 0, 0, 0, 0],
    });
    const sheet = evaluateTemplate(block, lone);
    expectRow(sheet, ROW.depreciation, [5, 10, 10, 10, 10], "dep");
    expectRow(sheet, ROW.netClosing, [95, 85, 75, 65, 55], "net");
  });

  it("depreciates faster on a shorter life", () => {
    // by hand, life 5: year one 200 / 5 + 30 / 5 / 2 = 40 + 3 = 43, leaving 187.
    const sheet = evaluateTemplate(block, { [LIFE]: 5 });
    expect(sheet.value("B14")).toBe(43);
    expect(sheet.value("B15")).toBe(187);
  });

  it("holds net PP&E at zero on a short life instead of writing off more than there is", () => {
    // by hand, life 3: 200 / 3 + 30 / 6 = 71.67, net 158.33; 230 / 3 + 33 / 6 =
    // 82.17, net 109.17; 263 / 3 + 36 / 6 = 93.67, net 51.5; year four would be
    // 299 / 3 + 38 / 6 = 106 but only 51.5 + 38 = 89.5 is left, net 0; year five
    // is capped at its own 40, net 0. Uncapped, net would end at -95.5.
    const sheet = evaluateTemplate(block, { [LIFE]: 3 });
    const expected = [200 / 3 + 5, 230 / 3 + 5.5, 263 / 3 + 6, 89.5, 40];
    expectRow(sheet, ROW.depreciation, expected, "dep");
    expectRow(
      sheet,
      ROW.netClosing,
      [158 + 1 / 3, 109 + 1 / 6, 51.5, 0, 0],
      "net",
    );
    for (const column of YEARS) {
      const net = sheet.value(addr(column, ROW.netClosing)) as number;
      expect(net, column).toBeGreaterThanOrEqual(0);
    }
  });

  it("writes a one-year life off in full and holds net at zero after it", () => {
    // by hand, life 1, capex 40 a year: year one 200 + 20 = 220 of the 240 there
    // is, net 20; year two 240 + 20 = 260 against 20 + 40 = 60, so 60 and net 0.
    const sheet = evaluateTemplate(
      block,
      typed({ ...DEMO, life: 1, capex: [40, 40, 40, 40, 40] }),
    );
    expectRow(sheet, ROW.depreciation, [220, 60, 40, 40, 40], "dep");
    expectRow(sheet, ROW.netClosing, [20, 0, 0, 0, 0], "net");
  });

  it("shows #DIV/0! through the depreciation and net for a life of zero or none, never a depreciation of nothing", () => {
    for (const life of [0, null]) {
      const sheet = evaluateTemplate(block, { [LIFE]: life });
      for (const column of YEARS) {
        const error = new ExcelError("#DIV/0!");
        expect(sheet.value(addr(column, ROW.depreciation)), column).toEqual(
          error,
        );
        expect(sheet.value(addr(column, ROW.netClosing)), column).toEqual(
          error,
        );
      }
      // Revenue, capex and the gross balance do not depend on the life.
      expect(sheet.value("B7")).toBeCloseTo(0.1, 12);
      expect(sheet.value("F10")).toBe(377);
    }
  });

  it("writes the depreciation and the rolls where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(13, 1)).toBe("=MIN(D16+D17,D13/D6+D17/D6/2)");
    expect(formula(14, 1)).toBe("=D16+D17-D18");
    expect(formula(9, 1)).toBe("=D13+D10");
    expect(formula(8, 2)).toBe("=D14");
    expect(formula(11, 2)).toBe("=D19");
    expect(formula(6, 1)).toBe('=IF(N(D9)>0,D10/D9,"n.m.")');
  });
});

describe("the capex and depreciation block in the picker", () => {
  it("is listed under Operating, carrying this very block", () => {
    const entry = entryById("template-capex-depreciation");
    expect(entry?.group).toBe("Operating");
    expect(entry?.block).toEqual(block);
  });

  it("is found by the words a modeller types for it", () => {
    for (const query of ["capex", "depreciation", "pp&e", "roll forward"]) {
      const ids = filterTemplates(TEMPLATES, query).map((entry) => entry.id);
      expect(ids, query).toContain("template-capex-depreciation");
    }
  });
});
