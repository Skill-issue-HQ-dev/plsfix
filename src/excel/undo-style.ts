// pls,fix Undo and the cell style: a style name goes back only to cells that wear another by now, and
// what that reset takes with it (lock, hidden formulas, rotation, shrink, strikethrough) is captured
// and written back on those same cells. Invariant: the style is written before the number format
// and the formats (src/excel/undo.ts orders them); a style the workbook lost goes to Normal.

/** One style name per captured cell; undefined where nothing is to be written. */
export type StyleGrid = (string | undefined)[][];

const NORMAL = "Normal";
const lower = (name: string): string => name.toLowerCase();

function nameOf(cell: Excel.CellProperties | undefined): string | undefined {
  return typeof cell?.style === "string" && cell.style !== ""
    ? cell.style
    : undefined;
}

/** The captured style of every cell that wears another one now. */
export function changedStyles(
  captured: Excel.CellProperties[][],
  current: Excel.CellProperties[][],
): StyleGrid {
  return captured.map((row, r) =>
    row.map((cell, c) => {
      const was = nameOf(cell);
      const now = nameOf(current[r]?.[c]);
      if (was === undefined) return undefined;
      return now !== undefined && lower(now) === lower(was) ? undefined : was;
    }),
  );
}

/**
 * The names that can really go back: one the workbook has stays, one it lost
 * becomes Normal (unless the cell already wears Normal). `known` holds the
 * workbook's style names in lower case.
 */
export function resolveStyles(
  grid: StyleGrid,
  current: Excel.CellProperties[][],
  known: ReadonlySet<string>,
): StyleGrid {
  return grid.map((row, r) =>
    row.map((name, c) => {
      if (name === undefined || known.has(lower(name))) return name;
      const now = nameOf(current[r]?.[c]);
      return now !== undefined && lower(now) === lower(NORMAL)
        ? undefined
        : NORMAL;
    }),
  );
}

export function anyStyle(grids: readonly StyleGrid[]): boolean {
  return grids.some((grid) =>
    grid.some((row) => row.some((name) => name !== undefined)),
  );
}

/** The grid setCellProperties takes: a style where there is one, nothing else. */
export function styleOnlyProperties(
  grid: StyleGrid,
): Excel.SettableCellProperties[][] {
  return grid.map((row) =>
    row.map((name) => (name === undefined ? {} : { style: name })),
  );
}

/**
 * Reads what each captured block wears now and, only when a style has to go
 * back, the workbook's style table: one sync for the first, a second for the
 * other. Returns, per block, the style names to write.
 */
export async function stylesToRestore(
  context: Excel.RequestContext,
  blocks: readonly { formats: Excel.CellProperties[][] }[],
  ranges: readonly Excel.Range[],
): Promise<StyleGrid[]> {
  const reads = ranges.map((range) => range.getCellProperties({ style: true }));
  await context.sync();
  const changed = blocks.map((block, index) =>
    changedStyles(block.formats, reads[index]!.value),
  );
  if (!anyStyle(changed)) return changed;

  const styles = context.workbook.styles;
  styles.load("items/name");
  await context.sync();
  const known = new Set(styles.items.map((style) => lower(style.name)));
  return changed.map((grid, index) =>
    resolveStyles(grid, reads[index]!.value, known),
  );
}

const isBoolean = (value: unknown): value is boolean =>
  typeof value === "boolean";

/**
 * The categories a style resets that the ordinary capture does not put back on
 * its own, read from one captured cell and rebuilt field by field: a boolean or
 * a number the host really answered, never an annotation, a null or a stray key.
 */
export function styleCategoryFormat(
  format: Excel.CellPropertiesFormat | undefined,
): Excel.CellPropertiesFormat {
  const out: Excel.CellPropertiesFormat = {};
  if (!format) return out;

  const strikethrough = format.font?.strikethrough;
  if (isBoolean(strikethrough)) out.font = { strikethrough };

  const protection: Excel.CellPropertiesProtection = {};
  const locked = format.protection?.locked;
  const formulaHidden = format.protection?.formulaHidden;
  if (isBoolean(locked)) protection.locked = locked;
  if (isBoolean(formulaHidden)) protection.formulaHidden = formulaHidden;
  if (Object.keys(protection).length > 0) out.protection = protection;

  const rotation = format.textOrientation;
  if (typeof rotation === "number" && Number.isFinite(rotation)) {
    out.textOrientation = rotation;
  }
  if (isBoolean(format.shrinkToFit)) out.shrinkToFit = format.shrinkToFit;
  return out;
}
