// The typed-over cases the LibreOffice cross-check runs on top of every block's
// demo values: [block id, what was typed, the cells overtyped by A1 address]. Owns
// the operating blocks' edge cases (a zero or blank life, text where revenue belongs,
// an out-of-range selector, overtyped links), where two engines most likely part ways.

import { type Computed, ExcelError } from "../formula-eval";

export type Overrides = Record<string, Computed>;

const NA = new ExcelError("#N/A");
const NO_CAPEX = { B6: 0, C6: 0, D6: 0, E6: 0, F6: 0 };
const SHRINKING = {
  B3: 500,
  B5: 12.5,
  C4: -0.1,
  D4: -0.05,
  E4: -0.05,
  F4: -0.02,
  C6: 0.01,
  D6: 0,
  E6: 0.01,
  F6: 0,
};
// Year one loses money and the opening sheet is rebuilt to balance: PP&E 50,
// working capital 5, cash 0, debt 20, equity 35.
const LOSS = {
  C3: 80,
  D3: 90,
  E3: 100,
  C4: 0.05,
  D4: 0.1,
  E4: 0.2,
  C6: 8,
  D6: 8,
  E6: 8,
  C8: 0.15,
  D8: 0.15,
  E8: 0.15,
  C24: 5,
  D24: 5,
  E24: 5,
  C25: 0,
  D25: 0,
  E25: 0,
  B15: 50,
  B16: 5,
  B17: 0,
  B18: 20,
  B19: 35,
};

export const LIBREOFFICE_CASES: [string, string, Overrides][] = [
  ["revenue-build", "zero volume", { B3: 0 }],
  ["revenue-build", "last year's volume gone", { F4: -1 }],
  ["revenue-build", "shrinking", SHRINKING],
  ["revenue-build", "text volume", { B3: "n.a." }],
  ["capex-depreciation", "life 5", { B2: 5 }],
  ["capex-depreciation", "life 3, net held at zero", { B2: 3 }],
  ["capex-depreciation", "life 1", { B2: 1 }],
  ["capex-depreciation", "life 0", { B2: 0 }],
  ["capex-depreciation", "blank life", { B2: null }],
  ["capex-depreciation", "text life", { B2: "x" }],
  ["capex-depreciation", "no revenue in year 3", { D5: 0 }],
  ["capex-depreciation", "text revenue in year 3", { D5: "n.a." }],
  ["capex-depreciation", "error in revenue", { D5: NA }],
  ["capex-depreciation", "net below gross", { B9: 350, B12: 200 }],
  ["capex-depreciation", "no capex", NO_CAPEX],
  ["three-statement", "a loss in year one", LOSS],
  ["three-statement", "opening equity out of balance", { B19: 45 }],
  ["three-statement", "year one cash typed over", { C17: 20 }],
  ["three-statement", "year two net debt typed over", { D20: 20 }],
  ["three-statement", "year two FCF typed over", { D26: 12 }],
  ["three-statement", "year two D&A add-back typed over", { D23: 5 }],
  ["three-statement", "cash and net debt typed over", { C17: 20, C20: 28 }],
  ["three-statement", "net cash to begin with", { B17: 50, B18: 0, B19: 130 }],
  ["scenario-switch", "selector 2", { B2: 2 }],
  ["scenario-switch", "selector 3", { B2: 3 }],
  ["scenario-switch", "selector 0", { B2: 0 }],
  ["scenario-switch", "selector 4", { B2: 4 }],
  ["scenario-switch", "selector 2.9", { B2: 2.9 }],
  ["scenario-switch", "selector -1", { B2: -1 }],
  ["scenario-switch", "text selector", { B2: "x" }],
  ["scenario-switch", "blank selector", { B2: null }],
];
