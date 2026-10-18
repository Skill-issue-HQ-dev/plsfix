// The WACC block against the finance it names: the beta relevered for the
// capital structure, CAPM with a size premium for the cost of equity, the tax
// shield on the cost of debt, capital weights off equity value and net debt.
// The formulas are computed by test/formula-eval.ts and held to the closed form
// written out independently below, on the demo case and on others typed over it.
import { describe, expect, it } from "vitest";
import {
  evaluateTemplate,
  ExcelError,
  type Value,
} from "../../test/formula-eval";
import { resolveFormula } from "../template-cells";
import { wacc } from "./wacc";

// Column index 2, row index 4: the corner the formula samples resolve against.
const C5 = { row: 4, column: 2 };

// The cell each input and each answer sits in, the block's corner being A1.
const CELL = {
  riskFree: "B2",
  betaUnlevered: "B3",
  debtToEquity: "B4",
  betaRelevered: "B5",
  premium: "B6",
  specific: "B7",
  costOfEquity: "B8",
  costOfDebt: "B10",
  taxRate: "B11",
  afterTax: "B12",
  equity: "B14",
  netDebt: "B15",
  debtInCapital: "B16",
  capital: "B17",
  equityWeight: "B18",
  debtWeight: "B19",
  wacc: "B21",
} as const;

interface Inputs {
  riskFree: number;
  betaUnlevered: number;
  premium: number;
  specific: number;
  costOfDebt: number;
  taxRate: number;
  equity: number;
  netDebt: number;
}

const DEMO: Inputs = {
  riskFree: 0.03,
  betaUnlevered: 0.95,
  premium: 0.055,
  specific: 0,
  costOfDebt: 0.06,
  taxRate: 0.2,
  equity: 750,
  netDebt: 250,
};

// The textbook, not the block's cells: a company with net cash has no debt in
// its capital, so debt is floored at zero; beta is relevered at the book's
// debt / equity and the cost of equity is CAPM plus the size premium.
function reference(inputs: Inputs) {
  const debt = Math.max(inputs.netDebt, 0);
  const betaRelevered =
    inputs.betaUnlevered * (1 + (1 - inputs.taxRate) * (debt / inputs.equity));
  const costOfEquity =
    inputs.riskFree + betaRelevered * inputs.premium + inputs.specific;
  const afterTax = inputs.costOfDebt * (1 - inputs.taxRate);
  const capital = inputs.equity + debt;
  return {
    betaRelevered,
    costOfEquity,
    wacc:
      (inputs.equity / capital) * costOfEquity + (debt / capital) * afterTax,
  };
}

function overrides(inputs: Inputs): Record<string, Value> {
  return {
    [CELL.riskFree]: inputs.riskFree,
    [CELL.betaUnlevered]: inputs.betaUnlevered,
    [CELL.premium]: inputs.premium,
    [CELL.specific]: inputs.specific,
    [CELL.costOfDebt]: inputs.costOfDebt,
    [CELL.taxRate]: inputs.taxRate,
    [CELL.equity]: inputs.equity,
    [CELL.netDebt]: inputs.netDebt,
  };
}

describe("the WACC block", () => {
  const block = wacc();
  const cell = (address: string) =>
    block.cells[Number(address.slice(1)) - 1]?.[1];

  it("is a 21 by 2 block named WACC, titled by its name", () => {
    expect(block.id).toBe("wacc");
    expect(block.name).toBe("WACC");
    expect([block.rows, block.cols]).toEqual([21, 2]);
    expect(block.cells[0]?.[0]).toMatchObject({ kind: "title", v: "WACC" });
  });

  it("labels each row for what it holds, units where there are any", () => {
    const labels = block.cells.map((row) => row[0]?.v);
    expect(labels).toEqual([
      "WACC",
      "Risk-free rate",
      "Beta (unlevered)",
      "Debt / equity",
      "Beta (relevered)",
      "Equity risk premium",
      "Size / specific premium",
      "Cost of equity (CAPM)",
      undefined,
      "Pre-tax cost of debt",
      "Tax rate",
      "After-tax cost of debt",
      undefined,
      "Equity value (EUR m)",
      "Net debt (EUR m)",
      "Debt in capital (EUR m)",
      "Capital (EUR m)",
      "Weight of equity",
      "Weight of debt",
      undefined,
      "WACC",
    ]);
  });

  it("types the eight inputs as inputs and everything derived as a formula", () => {
    const inputs = [
      CELL.riskFree,
      CELL.betaUnlevered,
      CELL.premium,
      CELL.specific,
      CELL.costOfDebt,
      CELL.taxRate,
      CELL.equity,
      CELL.netDebt,
    ];
    const derived = [
      CELL.debtToEquity,
      CELL.betaRelevered,
      CELL.costOfEquity,
      CELL.afterTax,
      CELL.debtInCapital,
      CELL.capital,
      CELL.equityWeight,
      CELL.debtWeight,
    ];
    for (const address of inputs) {
      expect(cell(address), address).toMatchObject({ kind: "input" });
      expect(cell(address)?.f, address).toBeUndefined();
    }
    for (const address of derived) {
      expect(cell(address), address).toMatchObject({ kind: "formula" });
      expect(cell(address)?.v, address).toBeUndefined();
    }
    expect(cell(CELL.wacc)).toMatchObject({
      kind: "result",
      format: "percent",
    });
  });

  it("shows the betas and the debt / equity to two decimals", () => {
    for (const address of [
      CELL.betaUnlevered,
      CELL.debtToEquity,
      CELL.betaRelevered,
    ]) {
      expect(cell(address)?.format, address).toBe("decimal2");
    }
    expect(cell(CELL.specific)?.format).toBe("percent");
  });

  // Worked out apart from the block: 0.95 x (1 + 0.8 x 250 / 750) = 1.2033;
  // 3.0 % + 1.2033 x 5.5 % = 9.618 %; 75 % x 9.618 % + 25 % x 4.8 % = 8.414 %.
  it("relevers the beta, prices equity with CAPM and weights off capital", () => {
    const sheet = evaluateTemplate(block);
    expect(sheet.value(CELL.debtToEquity)).toBeCloseTo(1 / 3, 12);
    expect(sheet.value(CELL.betaRelevered)).toBeCloseTo(1.2033, 4);
    expect(sheet.value(CELL.costOfEquity)).toBeCloseTo(0.09618, 5);
    expect(sheet.value(CELL.afterTax)).toBeCloseTo(0.048, 12);
    expect(sheet.value(CELL.debtInCapital)).toBe(250);
    expect(sheet.value(CELL.capital)).toBe(1000);
    expect(sheet.value(CELL.equityWeight)).toBeCloseTo(0.75, 12);
    expect(sheet.value(CELL.debtWeight)).toBeCloseTo(0.25, 12);
    expect(sheet.value(CELL.wacc)).toBeCloseTo(0.08414, 5);
  });

  it("adds a size premium to the cost of equity and nothing else", () => {
    const base = evaluateTemplate(block);
    const sized = evaluateTemplate(block, { [CELL.specific]: 0.02 });
    expect(sized.value(CELL.costOfEquity)).toBeCloseTo(
      (base.value(CELL.costOfEquity) as number) + 0.02,
      12,
    );
    expect(sized.value(CELL.betaRelevered)).toBe(
      base.value(CELL.betaRelevered),
    );
    expect(sized.value(CELL.wacc)).toBeCloseTo(
      (base.value(CELL.wacc) as number) + 0.75 * 0.02,
      12,
    );
  });

  it("relevers harder as the debt grows and not at all without it", () => {
    const geared = evaluateTemplate(block, { [CELL.netDebt]: 750 });
    expect(geared.value(CELL.debtToEquity)).toBe(1);
    expect(geared.value(CELL.betaRelevered)).toBeCloseTo(0.95 * 1.8, 12);
    const free = evaluateTemplate(block, { [CELL.netDebt]: 0 });
    expect(free.value(CELL.betaRelevered)).toBe(0.95);
  });

  it("weights add up to one whatever the capital structure", () => {
    const sheet = evaluateTemplate(block, {
      [CELL.equity]: 300,
      [CELL.netDebt]: 700,
    });
    const total =
      (sheet.value(CELL.equityWeight) as number) +
      (sheet.value(CELL.debtWeight) as number);
    expect(total).toBeCloseTo(1, 12);
  });

  it.each<[string, Inputs]>([
    ["the demo case", DEMO],
    [
      "a leveraged business",
      {
        ...DEMO,
        riskFree: 0.025,
        betaUnlevered: 1.1,
        premium: 0.06,
        costOfDebt: 0.085,
        taxRate: 0.15,
        equity: 400,
        netDebt: 600,
      },
    ],
    [
      "an unlevered one: equity only",
      { ...DEMO, betaUnlevered: 0.9, premium: 0.05, equity: 500, netDebt: 0 },
    ],
    [
      "a net cash company",
      { ...DEMO, betaUnlevered: 1.1, equity: 800, netDebt: -150 },
    ],
    [
      "no tax shield and a size premium",
      {
        ...DEMO,
        riskFree: 0.04,
        betaUnlevered: 1,
        premium: 0.05,
        specific: 0.015,
        costOfDebt: 0.07,
        taxRate: 0,
        equity: 600,
        netDebt: 400,
      },
    ],
  ])("agrees with the textbook on %s", (_name, inputs) => {
    const sheet = evaluateTemplate(block, overrides(inputs));
    const want = reference(inputs);
    expect(sheet.value(CELL.betaRelevered)).toBeCloseTo(want.betaRelevered, 12);
    expect(sheet.value(CELL.costOfEquity)).toBeCloseTo(want.costOfEquity, 12);
    expect(sheet.value(CELL.wacc)).toBeCloseTo(want.wacc, 12);
  });

  it("treats net cash as no debt: weights 100 and 0, beta not relevered", () => {
    const sheet = evaluateTemplate(block, { [CELL.netDebt]: -200 });
    expect(sheet.value(CELL.debtInCapital)).toBe(0);
    expect(sheet.value(CELL.capital)).toBe(750);
    expect(sheet.value(CELL.equityWeight)).toBe(1);
    expect(sheet.value(CELL.debtWeight)).toBe(0);
    expect(sheet.value(CELL.betaRelevered)).toBe(0.95);
    expect(sheet.value(CELL.wacc)).toBeCloseTo(
      sheet.value(CELL.costOfEquity) as number,
      12,
    );
  });

  it("collapses to the cost of equity when there is no net debt", () => {
    const sheet = evaluateTemplate(block, { [CELL.netDebt]: 0 });
    expect(sheet.value(CELL.wacc)).toBeCloseTo(
      sheet.value(CELL.costOfEquity) as number,
      12,
    );
  });

  it("shows the divide-by-zero of an empty capital base instead of a number", () => {
    const sheet = evaluateTemplate(block, {
      [CELL.equity]: 0,
      [CELL.netDebt]: 0,
    });
    expect(sheet.value(CELL.wacc)).toEqual(new ExcelError("#DIV/0!"));
  });

  it("writes the relevering, CAPM and weights where a banker reads them", () => {
    const formula = (row: number) => {
      const found = block.cells[row]?.[1];
      return found?.f === undefined ? undefined : resolveFormula(found.f, C5);
    };
    expect(formula(3)).toBe("=D20/D18");
    expect(formula(4)).toBe("=D7*(1+(1-D15)*D8)");
    expect(formula(7)).toBe("=D6+D9*D10+D11");
    expect(formula(11)).toBe("=D14*(1-D15)");
    expect(formula(15)).toBe("=MAX(D19,0)");
    expect(formula(16)).toBe("=D18+D20");
    expect(formula(17)).toBe("=D18/D21");
    expect(formula(18)).toBe("=D20/D21");
    expect(formula(20)).toBe("=D22*D12+D23*D16");
  });
});
