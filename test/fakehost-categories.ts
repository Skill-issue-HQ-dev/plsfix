// The fake host's per-cell categories a style can reset beside the ones fakehost.ts models:
// lock and hidden formulas, rotation, shrink-to-fit (strikethrough rides in the font).
// A read answers only what was asked for, a write changes only what it names.

import { defaultCell, type FakeCell } from "./fakehost";
import type { StyleBody } from "./fakehost-styles";

type Format = Record<string, unknown>;

/** The category fields of one cell that a getCellProperties request asked for. */
export function readCategories(cell: FakeCell, wanted: Format): Format {
  const format: Format = {};
  if (wanted.textOrientation) format.textOrientation = cell.textOrientation;
  if (wanted.shrinkToFit) format.shrinkToFit = cell.shrinkToFit;
  if (wanted.protection) {
    format.protection = {
      locked: cell.locked,
      formulaHidden: cell.formulaHidden,
    };
  }
  return format;
}

/** The category fields of one setCellProperties format entry, written to a cell. */
export function writeCategories(cell: FakeCell, format: Format): void {
  if (format.textOrientation !== undefined) {
    cell.textOrientation = Number(format.textOrientation);
  }
  if (format.shrinkToFit !== undefined) {
    cell.shrinkToFit = Boolean(format.shrinkToFit);
  }
  const protection = format.protection as Format | undefined;
  if (protection?.locked !== undefined)
    cell.locked = Boolean(protection.locked);
  if (protection?.formulaHidden !== undefined) {
    cell.formulaHidden = Boolean(protection.formulaHidden);
  }
}

/**
 * Excel for the web hands its own read-only annotation out with a protection
 * object and refuses the object whole when it comes back with it.
 */
export function hasAnnotatedProtection(
  grid: (Record<string, unknown> | null)[][],
): boolean {
  return grid.some((row) =>
    row.some((props) => {
      const format = props?.format as Format | undefined;
      const protection = format?.protection as Format | undefined;
      return protection !== undefined && "@odata.type" in protection;
    }),
  );
}

function clone<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

/**
 * What applying a style does to one cell: each category the style includes is
 * copied over the cell's own, the way Excel's Cell Styles gallery does. The
 * alignment category also resets vertical alignment, wrapping, indent, rotation
 * and shrink-to-fit; the font category its strikethrough with the rest of the
 * font; the protection category the lock and the hidden formulas.
 */
export function copyCategories(cell: FakeCell, body: StyleBody): void {
  const { include } = body;
  if (include.number) cell.numberFormat = body.numberFormat;
  if (include.font) cell.font = clone(body.font);
  if (include.patterns) cell.fill = clone(body.fill);
  if (include.border) cell.borders = clone(body.borders);
  if (include.alignment) {
    const fresh = defaultCell();
    cell.horizontalAlignment = body.horizontalAlignment;
    cell.verticalAlignment = fresh.verticalAlignment;
    cell.wrapText = fresh.wrapText;
    cell.indentLevel = fresh.indentLevel;
    cell.textOrientation = fresh.textOrientation;
    cell.shrinkToFit = fresh.shrinkToFit;
  }
  if (include.protection) {
    cell.locked = body.locked;
    cell.formulaHidden = body.formulaHidden;
  }
}
