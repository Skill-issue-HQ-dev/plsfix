// The football-field data block: five valuation methods with a low, a high and
// the midpoint, in EUR millions, laid out as the Football field button reads its
// selection. Invariant: the header row is the block's first row, since a title
// above it would be read as the header and the real one as a method.

import {
  build,
  calc,
  head,
  input,
  label,
  type Template,
  type TemplateCell,
} from "../template-cells";
import { span } from "./refs";

const LOW = 1;
const HIGH = 2;
const HEADER = ["Method (EUR m)", "Low", "High", "Midpoint"];

// EUR millions: ranges that overlap around 470 to 510, the way a football
// field's methods agree on a zone, each wider or narrower by how it is built.
const METHODS: [name: string, low: number, high: number][] = [
  ["Trading comps", 465, 555],
  ["Precedents", 400, 600],
  ["DCF", 450, 580],
  ["52-week range", 410, 510],
  ["Analyst targets", 470, 560],
];

function methodRow(
  [name, low, high]: [string, number, number],
  row: number,
): TemplateCell[] {
  return [
    label(name),
    input(low, "decimal"),
    input(high, "decimal"),
    calc(`=AVERAGE(${span(row, LOW, row, HIGH)})`, "decimal"),
  ];
}

export function footballField(): Template {
  return build(
    "football-field",
    "Football field data",
    "Five valuation methods with a low, a high and the midpoint, laid out for the Football field chart: select the block and press Football field.",
    [
      HEADER.map((name) => head(name)),
      ...METHODS.map((method, index) => methodRow(method, index + 1)),
    ],
  );
}
