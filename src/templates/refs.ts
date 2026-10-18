// Placeholder builders for the template grids: `at(row, column)` names the cell
// that far from the block's top-left corner in the `{r+dr,c+dc}` grammar and
// `span` names the range between two of them. Owns the spelling only; the
// grammar itself lives in src/template-cells.ts.

function part(axis: "r" | "c", offset: number): string {
  if (offset === 0) return axis;
  return `${axis}${offset > 0 ? "+" : "-"}${String(Math.abs(offset))}`;
}

/** The placeholder for the cell `row` rows and `column` columns from the corner. */
export function at(row: number, column: number): string {
  return `{${part("r", row)},${part("c", column)}}`;
}

/** The placeholder range between two cells, corner to corner. */
export function span(
  fromRow: number,
  fromColumn: number,
  toRow: number,
  toColumn: number,
): string {
  return `${at(fromRow, fromColumn)}:${at(toRow, toColumn)}`;
}
