// The template registry the Tools tab's picker lists: each entry carries the id
// its Insert button has as data-action (`template-` + the block's id), its
// group, name, one short line and the block. Owns the list and its order; the
// grids live in src/template-blocks.ts and one file per block beside this one.

import {
  dcf,
  debtSchedule,
  ebitdaBridge,
  npvIrr,
  sensitivity,
  workingCapital,
} from "../template-blocks";
import { type Template } from "../template-cells";
import { capTable } from "./cap-table";
import { capexDepreciation } from "./capex-depreciation";
import { debtPaydown } from "./debt-paydown";
import { footballField } from "./football-field";
import { precedents } from "./precedents";
import { returns } from "./returns";
import { revenueBuild } from "./revenue-build";
import { scenarioSwitch } from "./scenario-switch";
import { sourcesUses } from "./sources-uses";
import { threeStatement } from "./three-statement";
import { tradingComps } from "./trading-comps";
import { wacc } from "./wacc";

export type TemplateGroup = "Valuation" | "LBO" | "Operating" | "Analysis";

/** The groups in the order the picker shows them. */
export const TEMPLATE_GROUPS: readonly TemplateGroup[] = [
  "Valuation",
  "LBO",
  "Operating",
  "Analysis",
];

export interface TemplateEntry {
  /** The `data-action` of its Insert button: `template-` + the block's id. */
  id: string;
  group: TemplateGroup;
  name: string;
  /** One short line for the picker row; the block carries the long one. */
  description: string;
  block: Template;
}

function entry(
  group: TemplateGroup,
  block: Template,
  description: string,
): TemplateEntry {
  return {
    id: `template-${block.id}`,
    group,
    name: block.name,
    description,
    block,
  };
}

/**
 * The entries grouped as the picker lists them: by `TEMPLATE_GROUPS` order, the
 * authoring order kept inside each group (the sort is stable). A block can be
 * added anywhere in the list below and still land in its group.
 */
export function inPickerOrder(
  entries: readonly TemplateEntry[],
): TemplateEntry[] {
  const rank = (one: TemplateEntry): number =>
    TEMPLATE_GROUPS.indexOf(one.group);
  return [...entries].sort((a, b) => rank(a) - rank(b));
}

/** Every template, grouped in the order the picker lists them. */
export const TEMPLATES: TemplateEntry[] = inPickerOrder([
  entry(
    "Valuation",
    dcf(),
    "Five-year cash flows to enterprise value, Gordon terminal value.",
  ),
  entry(
    "Valuation",
    wacc(),
    "Relevered beta, CAPM cost of equity, after-tax debt, weights.",
  ),
  entry(
    "Valuation",
    tradingComps(),
    "Six peers' multiples, statistics and an implied EV range.",
  ),
  entry(
    "Valuation",
    precedents(),
    "Six deals' EV / EBITDA, statistics and an implied EV.",
  ),
  entry(
    "Valuation",
    footballField(),
    "Low, high and midpoint per method, ready for the chart.",
  ),
  entry(
    "Valuation",
    npvIrr(),
    "An outlay and eight cash flows into NPV, IRR and payback.",
  ),
  entry(
    "LBO",
    debtSchedule(),
    "Level-payment loan with interest and balance per period.",
  ),
  entry(
    "LBO",
    sourcesUses(),
    "Sources and uses that tie, entry EV, multiple and equity share.",
  ),
  entry(
    "LBO",
    debtPaydown(),
    "Five-year term loan: interest, amortisation, sweep, leverage.",
  ),
  entry(
    "LBO",
    returns(),
    "Exit equity, MOIC, IRR and MOIC one multiple step either way.",
  ),
  entry(
    "LBO",
    capTable(),
    "A funding round: post-money, price, new shares, holders' stakes.",
  ),
  entry("Operating", workingCapital(), "DSO, DIO, DPO and the cash cycle."),
  entry(
    "Operating",
    revenueBuild(),
    "Volume x price over five years, growth and CAGR in EUR m.",
  ),
  entry(
    "Operating",
    capexDepreciation(),
    "PP&E roll-forward: capex and depreciation over five years.",
  ),
  entry(
    "Operating",
    threeStatement(),
    "Income statement, balance sheet and cash flow over three years.",
  ),
  entry(
    "Operating",
    scenarioSwitch(),
    "Base, upside and downside assumptions behind one selector.",
  ),
  entry(
    "Analysis",
    sensitivity(),
    "A 5 by 5 grid around a base value and two driver steps.",
  ),
  entry(
    "Analysis",
    ebitdaBridge(),
    "Opening to closing EBITDA in five steps, for the waterfall.",
  ),
]);

/** The entry whose Insert button carries `id`, or null for a stranger. */
export function entryById(id: string): TemplateEntry | null {
  return TEMPLATES.find((candidate) => candidate.id === id) ?? null;
}
