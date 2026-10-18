// The year strip the multi-year blocks share: the column of each year and the
// header row that numbers them. Pure; year n sits in column n, shifted right by
// the lead columns before year one (an Opening balance), the labels in column 0,
// so a formula names the year before as column n - 1.

import { head, numberCell, type TemplateCell } from "../template-cells";

/** The columns of years 1 to `count`, the labels in column 0, `lead` columns between. */
export function yearColumns(count: number, lead = 0): number[] {
  return Array.from({ length: count }, (_unused, index) => index + 1 + lead);
}

/** "Year", the lead columns' words (text), then 1 to `count`, as header cells. */
export function yearHeader(count: number, lead: string[] = []): TemplateCell[] {
  return [
    head("Year"),
    ...lead.map(head),
    ...yearColumns(count).map((year) => numberCell(year, "header", "whole")),
  ];
}
