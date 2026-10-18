// The three-statement mini block held to the direct method on four businesses
// typed over the demo: the same figures worked out in a loop, never through the
// block's rows, and a balance check of zero in every column. Computed by
// test/formula-eval.ts; the loss year is also worked out by hand.
import { describe, expect, it } from "vitest";
import { evaluateTemplate, type Value } from "../../test/formula-eval";
import { threeStatement } from "./three-statement";

// With the block's corner at A1: the opening balances in B, years one to three
// in C to E.
const OPENING = "B";
const YEARS = ["C", "D", "E"] as const;
const ROW = {
  revenue: 3,
  margin: 4,
  da: 6,
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
  capex: 24,
  wcIncrease: 25,
  fcf: 26,
} as const;

const addr = (column: string, row: number): string => `${column}${String(row)}`;

interface Case {
  revenue: number[];
  margin: number[];
  da: number[];
  rate: number[];
  taxRate: number[];
  capex: number[];
  wcIncrease: number[];
  // The opening balances; equity is what balances them.
  ppe: number;
  wc: number;
  cash: number;
  debt: number;
}

const DEMO: Case = {
  revenue: [120, 132, 144],
  margin: [0.2, 0.21, 0.22],
  da: [6, 6.5, 7],
  rate: [0.075, 0.075, 0.075],
  taxRate: [0.2, 0.2, 0.2],
  capex: [8, 9, 9.5],
  wcIncrease: [2, 2.5, 2.5],
  ppe: 60,
  wc: 20,
  cash: 10,
  debt: 50,
};

// A business that loses money in year one: EBITDA 4 against D&A 8 and interest 3.
const LOSS: Case = {
  revenue: [80, 90, 100],
  margin: [0.05, 0.1, 0.2],
  da: [8, 8, 8],
  rate: [0.15, 0.15, 0.15],
  taxRate: [0.2, 0.2, 0.2],
  capex: [5, 5, 5],
  wcIncrease: [0, 0, 0],
  ppe: 50,
  wc: 5,
  cash: 0,
  debt: 20,
};

const CASES: [string, Case][] = [
  ["the demo case", DEMO],
  [
    "thin margins and heavy interest",
    {
      revenue: [100, 100, 100],
      margin: [0.1, 0.1, 0.12],
      da: [5, 5, 5],
      rate: [0.1, 0.1, 0.1],
      taxRate: [0.25, 0.25, 0.25],
      capex: [4, 4, 4],
      wcIncrease: [1, -1, 0.5],
      ppe: 80,
      wc: 15,
      cash: 5,
      debt: 60,
    },
  ],
  ["a loss in year one", LOSS],
  [
    "net cash to begin with, earning interest, and no tax",
    {
      revenue: [200, 210, 220],
      margin: [0.3, 0.3, 0.3],
      da: [10, 10, 10],
      rate: [0.03, 0.03, 0.03],
      taxRate: [0, 0, 0],
      capex: [12, 12, 12],
      wcIncrease: [3, 3, 3],
      ppe: 100,
      wc: 10,
      cash: 50,
      debt: 0,
    },
  ],
];

// A case as the cells a modeller overtypes.
function typed(c: Case): Record<string, Value> {
  const cells: Record<string, Value> = {
    [addr(OPENING, ROW.ppe)]: c.ppe,
    [addr(OPENING, ROW.wc)]: c.wc,
    [addr(OPENING, ROW.cash)]: c.cash,
    [addr(OPENING, ROW.debt)]: c.debt,
    [addr(OPENING, ROW.equity)]: c.ppe + c.wc + c.cash - c.debt,
  };
  YEARS.forEach((column, i) => {
    cells[addr(column, ROW.revenue)] = c.revenue[i] ?? null;
    cells[addr(column, ROW.margin)] = c.margin[i] ?? null;
    cells[addr(column, ROW.da)] = c.da[i] ?? null;
    cells[addr(column, ROW.rate)] = c.rate[i] ?? null;
    cells[addr(column, ROW.taxRate)] = c.taxRate[i] ?? null;
    cells[addr(column, ROW.capex)] = c.capex[i] ?? null;
    cells[addr(column, ROW.wcIncrease)] = c.wcIncrease[i] ?? null;
  });
  return cells;
}

// One year by the direct method, as the lines the block shows.
type Year = Record<
  | "interest"
  | "ebt"
  | "tax"
  | "netIncome"
  | "fcf"
  | "ppe"
  | "wc"
  | "cash"
  | "equity"
  | "netDebt",
  number
>;

// Interest on the net debt a year opens on, tax only on a profit, the cash left
// after interest, tax, capex and the working capital absorbed, and each balance
// moved by its own flow.
function yearsOf(c: Case): Year[] {
  let { ppe, wc, cash } = c;
  let equity = c.ppe + c.wc + c.cash - c.debt;
  return c.revenue.map((revenue, i) => {
    const ebitda = revenue * (c.margin[i] ?? 0);
    const interest = (c.rate[i] ?? 0) * (c.debt - cash);
    const ebt = ebitda - (c.da[i] ?? 0) - interest;
    const tax = Math.max(ebt, 0) * (c.taxRate[i] ?? 0);
    const spent = (c.capex[i] ?? 0) + (c.wcIncrease[i] ?? 0);
    const fcf = ebitda - interest - tax - spent;
    ppe += (c.capex[i] ?? 0) - (c.da[i] ?? 0);
    wc += c.wcIncrease[i] ?? 0;
    cash += fcf;
    equity += ebt - tax;
    const netDebt = c.debt - cash;
    const netIncome = ebt - tax;
    return {
      interest,
      ebt,
      tax,
      netIncome,
      fcf,
      ppe,
      wc,
      cash,
      equity,
      netDebt,
    };
  });
}

const block = threeStatement();

// The row each line of a year sits on.
const LINE_ROWS: Record<keyof Year, number> = {
  interest: ROW.interest,
  ebt: ROW.ebt,
  tax: ROW.tax,
  netIncome: ROW.netIncome,
  fcf: ROW.fcf,
  ppe: ROW.ppe,
  wc: ROW.wc,
  cash: ROW.cash,
  equity: ROW.equity,
  netDebt: ROW.netDebt,
};

describe("the three-statement mini against the direct method", () => {
  it.each(CASES)("agrees on %s", (_name, c) => {
    const sheet = evaluateTemplate(block, typed(c));
    yearsOf(c).forEach((year, i) => {
      const column = YEARS[i] as string;
      for (const [line, row] of Object.entries(LINE_ROWS)) {
        const want = year[line as keyof Year];
        expect(sheet.value(addr(column, row)), `${line} ${column}`).toBeCloseTo(
          want,
          9,
        );
      }
    });
    for (const column of [OPENING, ...YEARS]) {
      expect(sheet.value(addr(column, ROW.check)), `check ${column}`).toBe(0);
    }
  });

  it("charges no tax on a loss and carries the whole loss into net income and equity", () => {
    // by hand, year one of the loss case: EBITDA 80 x 5 % = 4; EBIT 4 - 8 = -4;
    // interest 15 % of 20 = 3; EBT -7; no tax, so net income is -7 and FCF 4 - 3
    // - 0 - 5 = -4; cash 0 - 4 = -4; net debt 20 + 4 = 24; equity 35 - 7 = 28.
    const sheet = evaluateTemplate(block, typed(LOSS));
    expect(sheet.value("C10")).toBe(-7);
    expect(sheet.value("C12")).toBe(0);
    expect(sheet.value("C13")).toBe(-7);
    expect(sheet.value("C26")).toBe(-4);
    expect(sheet.value("C17")).toBe(-4);
    expect(sheet.value("C20")).toBe(24);
    expect(sheet.value("C19")).toBe(28);
  });
});
