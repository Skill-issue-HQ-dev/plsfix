// The unused-style scrubber: every custom cell style in the workbook, matched
// against the styles the cells themselves are wearing. One sync per phase - the
// style table and the sheet list, then every used range's extent, then the
// grids of the sheets small enough to read - never one sync per sheet.

import { queueRegistryRead, registryJson } from "./custom-style-store";
import {
  pickScannableSheets,
  SCAN_CELL_CAP,
  SELECTION_CELL_CAP,
  SHEET_SCAN_CELL_CAP,
} from "./internal";
import { registeredNames } from "../custom-styles";
import { unusedStyles, type WorkbookStyle } from "../styles-audit";

export interface StyleScan {
  // The custom styles no cell wears, sorted.
  unused: string[];
  // Every style in the workbook, built-ins included: what the count is out of.
  total: number;
  // Sheets whose used range was too large to read. A style worn only there
  // would look unused, so the pane says the answer is incomplete and the
  // delete refuses to run.
  skippedSheets: string[];
}

type CellGrid = OfficeExtension.ClientResult<Excel.CellProperties[][]>;

// Which sheets are worth reading is the same question Super Find and Prepare
// for sharing ask, so it is asked in the same place: an empty sheet dresses no
// cells, and a sheet - or a run of them - past the cap would overflow the
// request payload, so it is reported as skipped rather than half-read. Style
// properties are the heaviest read of the three, which is why the cap matters
// most here.
function readStyles(
  sheets: Excel.Worksheet[],
  ranges: Excel.Range[],
  cap: number,
  totalCap: number,
): { grids: CellGrid[]; skippedSheets: string[] } {
  const { scanned, skippedSheets } = pickScannableSheets(
    sheets,
    ranges,
    cap,
    totalCap,
  );
  return {
    grids: scanned.map((sheet) =>
      sheet.range.getCellProperties({ style: true }),
    ),
    skippedSheets,
  };
}

function wornStyles(grids: CellGrid[]): Set<string> {
  const worn = new Set<string>();
  for (const grid of grids) {
    for (const row of grid.value) {
      for (const cell of row) {
        if (typeof cell.style === "string") worn.add(cell.style);
      }
    }
  }
  return worn;
}

// Hidden sheets are read as well: a style worn only on a back-room sheet is in
// use, and deleting it would restyle those cells behind the modeller's back.
async function scanStyles(
  context: Excel.RequestContext,
  cap: number,
  totalCap: number,
): Promise<StyleScan> {
  const styles = context.workbook.styles;
  styles.load("items/name,items/builtIn");
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  // The styles pls,fix saved ride the same batch: they are never "unused".
  const registry = queueRegistryRead(context);
  await context.sync();

  const all: WorkbookStyle[] = styles.items.map((style) => ({
    name: style.name,
    builtIn: style.builtIn,
  }));
  const ranges = sheets.items.map((sheet) =>
    sheet.getUsedRangeOrNullObject(true),
  );
  for (const range of ranges) range.load("isNullObject,cellCount");
  await context.sync();

  const { grids, skippedSheets } = readStyles(
    sheets.items,
    ranges,
    cap,
    totalCap,
  );
  await context.sync();

  return {
    unused: unusedStyles(
      all,
      wornStyles(grids),
      registeredNames(registryJson(registry)),
    ),
    total: all.length,
    skippedSheets,
  };
}

function wearerCounts(grids: CellGrid[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const grid of grids) {
    for (const row of grid.value) {
      for (const cell of row) {
        if (typeof cell.style !== "string") continue;
        const key = cell.style.toLowerCase();
        counts.set(key, (counts.get(key) ?? 0) + 1);
      }
    }
  }
  return counts;
}

export interface StyleWearers {
  // Cells wearing each style, keyed by the name in lower case.
  counts: Map<string, number>;
  // Sheets too large to read: the counts leave out whatever they hold.
  skippedSheets: string[];
}

// How many cells wear each style, read the way the scrubber reads (hidden sheets
// included, sheets over the caps skipped and named) but under the caps Find and the
// model check use: custom styles count the cells a delete would send back to
// Normal with it, and a real model's big sheets have to be countable.
export async function scanStyleWearers(
  context: Excel.RequestContext,
  cap = SHEET_SCAN_CELL_CAP,
  totalCap = SHEET_SCAN_CELL_CAP,
): Promise<StyleWearers> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const ranges = sheets.items.map((sheet) =>
    sheet.getUsedRangeOrNullObject(true),
  );
  for (const range of ranges) range.load("isNullObject,cellCount");
  await context.sync();

  const { grids, skippedSheets } = readStyles(
    sheets.items,
    ranges,
    cap,
    totalCap,
  );
  await context.sync();
  return { counts: wearerCounts(grids), skippedSheets };
}

export async function listUnusedStyles(
  maxCells = SELECTION_CELL_CAP,
  maxTotalCells = SCAN_CELL_CAP,
): Promise<StyleScan> {
  return Excel.run((context) => scanStyles(context, maxCells, maxTotalCells));
}

// Irreversible, and every cell wearing a deleted style is restyled with it, so
// the list is re-derived here rather than trusted from the pane; a workbook
// with a sheet too large to read is refused outright, because a style could
// well be in use on it.
export async function deleteUnusedStyles(
  names: string[],
  maxCells = SELECTION_CELL_CAP,
  maxTotalCells = SCAN_CELL_CAP,
): Promise<number> {
  return Excel.run(async (context) => {
    const scan = await scanStyles(context, maxCells, maxTotalCells);
    if (scan.skippedSheets.length > 0) {
      throw new Error("styles: some sheets were too large to scan");
    }

    const unused = new Set(scan.unused);
    const doomed = names.filter((name) => unused.has(name));
    const styles = context.workbook.styles;
    for (const name of doomed) styles.getItem(name).delete();
    await context.sync();
    return doomed.length;
  });
}
