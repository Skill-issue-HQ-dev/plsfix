// The scenario-switch block against what a selector must do: one typed number
// picks Base, Upside or Downside for four assumptions through CHOOSE, and a
// three-year projection reads only the live column. Computed by
// test/formula-eval.ts and held to a closed form worked out independently.
import { describe, expect, it } from "vitest";
import { cellAt, errorsIn } from "../../test/block-grid";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { entryById, TEMPLATES } from "./index";
import { scenarioSwitch } from "./scenario-switch";
import { filterTemplates } from "./search";

// Column index 2, row index 4: the corner the pinned formulas resolve against.
const C5 = { row: 4, column: 2 };

// With the block's corner at A1: the selector in B2, the Base, Upside and
// Downside columns B to D with Live in E; the base year in B and years one to
// three in C to E.
const SELECTOR = "B2";
const SCENARIO_COLUMNS = ["B", "C", "D"] as const;
const YEAR_COLUMNS = ["C", "D", "E"] as const;
const LIVE = "E";
const ROW = {
  growth: 6,
  margin: 7,
  capex: 8,
  tax: 9,
  revenue: 12,
  ebitda: 13,
  capexSpend: 14,
  taxPaid: 15,
  fcf: 16,
} as const;
const ASSUMPTION_ROWS = [ROW.growth, ROW.margin, ROW.capex, ROW.tax];
const PROJECTION_ROWS = [
  ROW.revenue,
  ROW.ebitda,
  ROW.capexSpend,
  ROW.taxPaid,
  ROW.fcf,
];
const BASE_REVENUE = 100;

const addr = (column: string, row: number): string => `${column}${String(row)}`;

interface Assumptions {
  growth: number;
  margin: number;
  capex: number;
  tax: number;
}

// Base, Upside, Downside, in the order the selector counts them.
const SCENARIOS: Assumptions[] = [
  { growth: 0.06, margin: 0.2, capex: 0.08, tax: 0.2 },
  { growth: 0.1, margin: 0.23, capex: 0.07, tax: 0.18 },
  { growth: 0.02, margin: 0.16, capex: 0.1, tax: 0.22 },
];
const NAMES = ["Base", "Upside", "Downside"];

// A year of the projection in closed form: revenue grown by compounding the live
// rate, EBITDA and capex off it, and tax charged only on a positive EBITDA less
// capex, the D&A that would shield it taken to be the capex.
function projectionOf(a: Assumptions, base: number, year: number) {
  const revenue = base * (1 + a.growth) ** year;
  const ebitda = revenue * a.margin;
  const capex = revenue * a.capex;
  const tax = Math.max(ebitda - capex, 0) * a.tax;
  return { revenue, ebitda, capex, tax, fcf: ebitda - capex - tax };
}

const block = scenarioSwitch();

describe("the scenario switch block", () => {
  it("is a 16 by 5 block named Scenario switch, titled by its name", () => {
    expect(block.id).toBe("scenario-switch");
    expect(block.name).toBe("Scenario switch");
    expect([block.rows, block.cols]).toEqual([16, 5]);
    expect(block.cells[0]?.[0]).toMatchObject({
      kind: "title",
      v: "Scenario switch",
    });
  });

  it("types the selector in as 1 and explains what 1, 2 and 3 mean beside it", () => {
    expect(cellAt(block, SELECTOR)).toMatchObject({
      kind: "input",
      format: "whole",
      v: 1,
    });
    expect(String(cellAt(block, "A2")?.v)).toMatch(/scenario/i);
    expect(cellAt(block, "A3")?.v).toBe("1 = Base, 2 = Upside, 3 = Downside");
    // Nothing to the right of the legend, so its text runs on across the row.
    for (const column of ["B", "C", "D", "E"]) {
      const cell = cellAt(block, addr(column, 3));
      expect(cell?.v, column).toBeUndefined();
      expect(cell?.f, column).toBeUndefined();
    }
  });

  it("heads the assumption columns Base, Upside, Downside and Live", () => {
    expect(
      ["A", "B", "C", "D", "E"].map((column) => cellAt(block, addr(column, 5))),
    ).toMatchObject([
      { kind: "header", v: "Assumptions" },
      { kind: "header", v: "Base" },
      { kind: "header", v: "Upside" },
      { kind: "header", v: "Downside" },
      { kind: "header", v: "Live" },
    ]);
  });

  it("types the three scenarios in and builds the Live column with CHOOSE", () => {
    SCENARIOS.forEach((scenario, index) => {
      const column = SCENARIO_COLUMNS[index] as string;
      const typed = [
        [ROW.growth, scenario.growth],
        [ROW.margin, scenario.margin],
        [ROW.capex, scenario.capex],
        [ROW.tax, scenario.tax],
      ] as const;
      for (const [row, value] of typed) {
        const cell = cellAt(block, addr(column, row));
        expect(cell, addr(column, row)).toMatchObject({
          kind: "input",
          format: "percent",
          v: value,
        });
      }
    });
    for (const row of ASSUMPTION_ROWS) {
      const live = cellAt(block, addr(LIVE, row));
      expect(live, addr(LIVE, row)).toMatchObject({
        kind: "formula",
        format: "percent",
      });
      expect(live?.f, addr(LIVE, row)).toMatch(/^=CHOOSE\(/);
    }
  });

  it("names the live scenario beside the selector", () => {
    NAMES.forEach((name, index) => {
      const sheet = evaluateTemplate(block, { [SELECTOR]: index + 1 });
      expect(sheet.value("C2"), name).toBe(name);
    });
  });

  it.each([1, 2, 3])(
    "carries scenario %i's four assumptions into the Live column",
    (selector) => {
      const sheet = evaluateTemplate(block, { [SELECTOR]: selector });
      const want = SCENARIOS[selector - 1] as Assumptions;
      expect(sheet.value(addr(LIVE, ROW.growth))).toBe(want.growth);
      expect(sheet.value(addr(LIVE, ROW.margin))).toBe(want.margin);
      expect(sheet.value(addr(LIVE, ROW.capex))).toBe(want.capex);
      expect(sheet.value(addr(LIVE, ROW.tax))).toBe(want.tax);
    },
  );

  it("takes the whole part of a fractional selector, as CHOOSE does", () => {
    const sheet = evaluateTemplate(block, { [SELECTOR]: 2.9 });
    expect(sheet.value("C2")).toBe("Upside");
    expect(sheet.value(addr(LIVE, ROW.growth))).toBe(0.1);
  });

  it.each<[string, Value]>([
    ["zero", 0],
    ["a negative", -1],
    ["past the last scenario", 4],
    ["text", "base"],
    ["blank", null],
  ])(
    "shows #VALUE! through the Live column and the projection for a selector of %s, and nowhere else",
    (_name, selector) => {
      const error = new ExcelError("#VALUE!");
      const sheet = evaluateTemplate(block, { [SELECTOR]: selector });
      expect(sheet.value("C2")).toEqual(error);
      for (const row of ASSUMPTION_ROWS) {
        expect(sheet.value(addr(LIVE, row)), `live ${String(row)}`).toEqual(
          error,
        );
      }
      for (const row of PROJECTION_ROWS) {
        for (const column of YEAR_COLUMNS) {
          expect(sheet.value(addr(column, row)), addr(column, row)).toEqual(
            error,
          );
        }
      }
      // The typed base year and the typed scenarios stay what they were.
      expect(sheet.value(addr("B", ROW.revenue))).toBe(BASE_REVENUE);
      expect(sheet.value(addr("D", ROW.margin))).toBe(0.16);
      // 1 readout, 4 live cells and 5 projection rows of 3 years: 20 errors.
      expect(errorsIn(block, { [SELECTOR]: selector })).toHaveLength(20);
    },
  );

  it.each([1, 2, 3])(
    "projects three years off scenario %i in closed form",
    (selector) => {
      const sheet = evaluateTemplate(block, { [SELECTOR]: selector });
      const scenario = SCENARIOS[selector - 1] as Assumptions;
      YEAR_COLUMNS.forEach((column, index) => {
        const year = projectionOf(scenario, BASE_REVENUE, index + 1);
        const checks: [number, number][] = [
          [ROW.revenue, year.revenue],
          [ROW.ebitda, year.ebitda],
          [ROW.capexSpend, year.capex],
          [ROW.taxPaid, year.tax],
          [ROW.fcf, year.fcf],
        ];
        for (const [row, want] of checks) {
          expect(sheet.value(addr(column, row)), addr(column, row)).toBeCloseTo(
            want,
            9,
          );
        }
      });
    },
  );

  it("works the three scenarios' first years by hand", () => {
    // by hand, Base: revenue 100 x 1.06 = 106; EBITDA 20 % = 21.2; capex 8 % =
    // 8.48; tax 20 % of (21.2 - 8.48 = 12.72) = 2.544; FCF 12.72 - 2.544 = 10.176.
    // Upside: 110; 25.3; 7.7; tax 18 % of 17.6 = 3.168; FCF 14.432. Downside:
    // 102; 16.32; 10.2; tax 22 % of 6.12 = 1.3464; FCF 4.7736.
    const want: [number, number[]][] = [
      [1, [106, 21.2, 8.48, 2.544, 10.176]],
      [2, [110, 25.3, 7.7, 3.168, 14.432]],
      [3, [102, 16.32, 10.2, 1.3464, 4.7736]],
    ];
    for (const [selector, figures] of want) {
      const sheet = evaluateTemplate(block, { [SELECTOR]: selector });
      PROJECTION_ROWS.forEach((row, index) => {
        expect(
          sheet.value(addr("C", row)),
          `scenario ${String(selector)} row ${String(row)}`,
        ).toBeCloseTo(figures[index] as number, 9);
      });
    }
    // by hand, Base year three: 100 x 1.06^3 = 119.1016; FCF = (23.82032 -
    // 9.528128) x 0.8 = 11.4337536.
    const base = evaluateTemplate(block);
    expect(base.value("E12")).toBeCloseTo(119.1016, 9);
    expect(base.value("E16")).toBeCloseTo(11.4337536, 9);
  });

  it("moves with a scenario's inputs only while that scenario is live", () => {
    // Upside growth typed as 50 %: Base is selected, so nothing moves.
    const quiet = evaluateTemplate(block, { C6: 0.5 });
    const demo = evaluateTemplate(block);
    for (const column of YEAR_COLUMNS) {
      expect(quiet.value(addr(column, ROW.revenue))).toBe(
        demo.value(addr(column, ROW.revenue)),
      );
    }
    // Select Upside and year one is 100 x 1.5.
    const live = evaluateTemplate(block, { C6: 0.5, [SELECTOR]: 2 });
    expect(live.value("C12")).toBe(150);
  });

  it("charges no tax in a year whose EBITDA does not cover its capex", () => {
    // by hand: Downside margin typed 5 % against capex of 10 %: year one revenue
    // 102, EBITDA 5.1, capex 10.2, EBITDA less capex -5.1; no tax, so FCF -5.1.
    const sheet = evaluateTemplate(block, { [SELECTOR]: 3, D7: 0.05 });
    expect(sheet.value("C13")).toBeCloseTo(5.1, 9);
    expect(sheet.value("C14")).toBeCloseTo(10.2, 9);
    expect(sheet.value("C15")).toBe(0);
    expect(sheet.value("C16")).toBeCloseTo(-5.1, 9);
  });

  it("types only the base year's revenue and leaves the rest of that column empty", () => {
    expect(cellAt(block, addr("B", ROW.revenue))).toMatchObject({
      kind: "input",
      format: "decimal",
      v: BASE_REVENUE,
    });
    for (const row of [ROW.ebitda, ROW.capexSpend, ROW.taxPaid, ROW.fcf]) {
      const cell = cellAt(block, addr("B", row));
      expect(cell?.v, `B${String(row)}`).toBeUndefined();
      expect(cell?.f, `B${String(row)}`).toBeUndefined();
    }
    for (const row of PROJECTION_ROWS) {
      for (const column of YEAR_COLUMNS) {
        expect(
          cellAt(block, addr(column, row)),
          addr(column, row),
        ).toMatchObject({ format: "decimal" });
      }
    }
  });

  it("computes with no error and no circular reference under each scenario", () => {
    for (const selector of [1, 2, 3]) {
      expect(errorsIn(block, { [SELECTOR]: selector }), `${selector}`).toEqual(
        [],
      );
    }
  });

  it("writes the CHOOSE and the projection where a banker audits them", () => {
    const formula = (row: number, column: number): string | undefined => {
      const cell = block.cells[row]?.[column];
      return cell?.f === undefined ? undefined : resolveFormula(cell.f, C5);
    };
    expect(formula(5, 4)).toBe("=CHOOSE(D6,D10,E10,F10)");
    expect(formula(11, 2)).toBe("=D16*(1+G10)");
    expect(formula(14, 2)).toBe("=MAX(E17-E18,0)*G13");
  });
});

describe("the scenario switch's free cash flow", () => {
  it("says it leaves working capital out, in its label and its description", () => {
    // EBITDA less capex less tax: no change in working capital is in the line.
    const name = String(cellAt(block, `A${String(ROW.fcf)}`)?.v);
    expect(name).toBe("Unlevered FCF (EUR m, no working capital)");
    expect(block.description).toContain("no working capital");
  });
});

describe("the scenario switch in the picker", () => {
  it("is listed under Operating, carrying this very block", () => {
    const entry = entryById("template-scenario-switch");
    expect(entry?.group).toBe("Operating");
    expect(entry?.block).toEqual(block);
  });

  it("is found by the words a modeller types for it", () => {
    for (const query of ["scenario", "base", "upside", "downside"]) {
      const ids = filterTemplates(TEMPLATES, query).map((entry) => entry.id);
      expect(ids, query).toContain("template-scenario-switch");
    }
  });
});
