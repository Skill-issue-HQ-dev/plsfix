// The WACC block: an unlevered beta relevered at the capital structure, a CAPM
// cost of equity with a size premium, the after-tax cost of debt and the capital
// weights into one weighted average cost of capital, in EUR millions. Owns the
// grid and its demo case; the invariant is that every placeholder names a cell
// inside the 21 by 2 block, and that net cash is no debt.

import {
  blank,
  blankRow,
  build,
  calc,
  input,
  label,
  result,
  type Template,
  type TemplateCell,
  title,
  total,
} from "../template-cells";
import { at } from "./refs";

const COLUMNS = 2;
// The value column: labels sit in column 0, every number in column 1.
const VALUE = 1;

// Where each line sits, so a formula reads by what it names.
const ROW = {
  riskFree: 1,
  betaUnlevered: 2,
  debtToEquity: 3,
  betaRelevered: 4,
  premium: 5,
  specific: 6,
  costOfEquity: 7,
  costOfDebt: 9,
  taxRate: 10,
  afterTax: 11,
  equity: 13,
  netDebt: 14,
  debtInCapital: 15,
  capital: 16,
  equityWeight: 17,
  debtWeight: 18,
} as const;

const value = (row: number): string => at(row, VALUE);

function line(name: string, cell: TemplateCell): TemplateCell[] {
  return [label(name), cell];
}

// Equity is riskier the more debt sits ahead of it: the unlevered beta is
// relevered at the debt / equity of the capital below, at the tax rate that
// shields the interest (Hamada). CAPM then adds the equity risk premium times
// that beta, plus a size or company-specific premium.
function costOfEquity(): TemplateCell[][] {
  const betaUnlevered = value(ROW.betaUnlevered);
  const gearing = value(ROW.debtToEquity);
  return [
    line("Risk-free rate", input(0.03, "percent")),
    line("Beta (unlevered)", input(0.95, "decimal2")),
    line(
      "Debt / equity",
      calc(`=${value(ROW.debtInCapital)}/${value(ROW.equity)}`, "decimal2"),
    ),
    line(
      "Beta (relevered)",
      calc(
        `=${betaUnlevered}*(1+(1-${value(ROW.taxRate)})*${gearing})`,
        "decimal2",
      ),
    ),
    line("Equity risk premium", input(0.055, "percent")),
    line("Size / specific premium", input(0, "percent")),
    line(
      "Cost of equity (CAPM)",
      calc(
        `=${value(ROW.riskFree)}+${value(ROW.betaRelevered)}*${value(ROW.premium)}` +
          `+${value(ROW.specific)}`,
        "percent",
      ),
    ),
  ];
}

// Interest is deductible, so debt costs its rate less the tax it saves.
function costOfDebt(): TemplateCell[][] {
  return [
    line("Pre-tax cost of debt", input(0.06, "percent")),
    line("Tax rate", input(0.2, "percent")),
    line(
      "After-tax cost of debt",
      calc(`=${value(ROW.costOfDebt)}*(1-${value(ROW.taxRate)})`, "percent"),
    ),
  ];
}

// Weights come off capital: equity value plus the debt in it. A company with
// net cash has no debt in its capital, so net debt is floored at zero, shown on
// its own row rather than buried in the weights.
function weights(): TemplateCell[][] {
  const debt = value(ROW.debtInCapital);
  return [
    line("Equity value (EUR m)", input(750, "decimal")),
    line("Net debt (EUR m)", input(250, "decimal")),
    line(
      "Debt in capital (EUR m)",
      calc(`=MAX(${value(ROW.netDebt)},0)`, "decimal"),
    ),
    line("Capital (EUR m)", calc(`=${value(ROW.equity)}+${debt}`, "decimal")),
    line(
      "Weight of equity",
      calc(`=${value(ROW.equity)}/${value(ROW.capital)}`, "percent"),
    ),
    line("Weight of debt", calc(`=${debt}/${value(ROW.capital)}`, "percent")),
  ];
}

export function wacc(): Template {
  const answer =
    `=${value(ROW.equityWeight)}*${value(ROW.costOfEquity)}` +
    `+${value(ROW.debtWeight)}*${value(ROW.afterTax)}`;
  return build(
    "wacc",
    "WACC",
    "An unlevered beta relevered at the capital structure, CAPM with a size premium into a cost of equity, a cost of debt and tax rate into its after-tax cost, equity value and net debt (floored at zero) into weights, and the WACC they make.",
    [
      [title("WACC"), blank("title")],
      ...costOfEquity(),
      blankRow(COLUMNS),
      ...costOfDebt(),
      blankRow(COLUMNS),
      ...weights(),
      blankRow(COLUMNS),
      [result("WACC"), total(answer, "percent")],
    ],
  );
}
