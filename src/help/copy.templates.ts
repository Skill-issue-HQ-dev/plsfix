// The Templates section's help copy: one line on what the picker is for, then
// one sentence per Insert button, keyed by its data-action (template-<block id>).
// Owns the whole "templates-heading" section, so a block adds its sentence here
// and nowhere else; copy.test.ts holds the sentences to the registry.

import type { HelpCopy } from "./copy";

export const TEMPLATE_HELP: HelpCopy = {
  "templates-heading": {
    about:
      "Search ready calculation blocks and write one at the active cell: blue inputs, black formulas, your styles; Undo takes it back.",
    buttons: {
      "template-dcf":
        "Writes a DCF block: five free cash flows, WACC and terminal growth in, discount factors, terminal value and enterprise value out.",
      "template-wacc":
        "Writes a WACC block: risk-free rate, unlevered beta, premium, cost of debt, tax and capital in, relevered beta, cost of equity and WACC out.",
      "template-trading-comps":
        "Writes a trading comps table: six peers in, three multiples, six statistics and an implied EV range out. Comps stats is for your own table.",
      "template-precedents":
        "Writes a precedents table: six deals in, EV / EBITDA, statistics, implied EV at 25th, median, 75th out. Comps stats is for your own table.",
      "template-football-field":
        "Writes the method, low, high and midpoint table the Football field button reads: keep three empty columns to its right, then press it.",
      "template-npv-irr":
        "Writes an NPV / IRR block: an outlay and eight cash flows in, NPV at your rate, IRR and payback out.",
      "template-debt-schedule":
        "Writes an annuity debt schedule: principal, rate and term in, payment, interest and balance per period out.",
      "template-sources-uses":
        "Writes sources and uses: debt, equity, uses and fees in, totals, a tie-out check, entry EV, EV / EBITDA and equity share out.",
      "template-debt-paydown":
        "Writes a five-year debt paydown: loan, rate, amortisation, cash for debt service and sweep % in, interest, balances and leverage out.",
      "template-returns":
        "Writes a returns block: entry equity, exit EBITDA, multiple and net debt in, exit equity, MOIC, IRR and a multiple sensitivity out.",
      "template-cap-table":
        "Writes a cap table: valuation, new money and option pool in, post-money, price, new shares and holdings before and after out.",
      "template-working-capital":
        "Writes working-capital days: revenue, COGS and balances in, DSO, DIO, DPO and the cash conversion cycle out.",
      "template-sensitivity":
        "Writes a two-way sensitivity grid: a base value and two drivers with steps in, a 5 by 5 grid of outcomes out.",
      "template-ebitda-bridge":
        "Writes an EBITDA bridge table shaped for the Waterfall button: opening, five steps, closing and a check row.",
      "template-revenue-build":
        "Writes a revenue build: year-one volume and price with yearly growth rates in, revenue, its growth and the year 1 to 5 CAGR out.",
      "template-capex-depreciation":
        "Writes a capex schedule: revenue, capex, life and opening gross and net PP&E in, straight-line depreciation and closing net PP&E out.",
      "template-three-statement":
        "Writes a three-statement mini: opening balances, revenue, margin, D&A, rates, capex and working capital in, a balanced sheet and FCF out.",
      "template-scenario-switch":
        "Writes a scenario switch: Base, Upside and Downside assumptions and a 1, 2 or 3 selector in, a Live column and three years of FCF out.",
    },
  },
};
