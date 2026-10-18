// Custom styles as real Excel cell styles (ExcelApi 1.7, Range.style): save the active cell's look under
// a name, put it on the selection, update it (every wearer follows), delete it, add the eight starters,
// and list them; they show in Home > Cell Styles and travel with the file. Which styles are pls,fix's
// lives in the setting plsfix.styles.v1 (custom-style-store.ts). Every flow ends on syncWrite.

import { countText } from "../ui/count-text";
import {
  checkStyleName,
  fromPaintSlot,
  isRegistered,
  onlyExisting,
  sameName,
  starterStyles,
  withoutStyle,
  withStyle,
  type RegistryEntry,
  type StylePreview,
} from "../custom-styles";
import { getActiveSettings } from "../settings";
import { startersDone, yoursNow } from "../starter-receipts";
import { selectedAreas } from "./areas";
import { putStyleOnAreas } from "./custom-style-put";
import {
  queueRegistryRead,
  queueRegistryWrite,
  registryEntries,
  registryOf,
} from "./custom-style-store";
import { writeStyle } from "./custom-style-write";
import { readFill } from "./fill-store";
import { SHEET_SCAN_CELL_CAP } from "./internal";
import { readActiveCell } from "./paintbrush";
import { syncWrite } from "./protection";
import { scanStyleWearers } from "./styles";

const STAGE_SAVE = "Save style";
const STAGE_APPLY = "Style";
const STAGE_UPDATE = "Update style";
const STAGE_DELETE = "Delete style";
const STAGE_STARTERS = "Starter styles";

// The style by name, queued with what a flow needs to know about it; whether it
// is there at all is settled by the sync the caller runs next.
function probe(context: Excel.RequestContext, name: string): Excel.Style {
  const style = context.workbook.styles.getItemOrNullObject(name);
  style.load("isNullObject,name,builtIn");
  return style;
}

const gone = (name: string): string =>
  `The style ${name} is no longer in this workbook.`;

// A style deleted through Excel's own Cell Styles is dropped from the list the
// moment a flow meets the gap, not left to show as a ghost.
async function forget(
  context: Excel.RequestContext,
  registry: Excel.Setting,
  name: string,
  stage: string,
): Promise<void> {
  const read = registryOf(registry);
  if (!read.ok || !isRegistered(read.entries, name)) return;
  queueRegistryWrite(context, withoutStyle(read.entries, name));
  await syncWrite(context, stage);
}

// ---------------------------------------------------------------------------
// Save, apply, update
// ---------------------------------------------------------------------------

/** Saves the active cell's look as a style named `rawName` and lists it. */
export async function saveCustomStyle(rawName: string): Promise<string> {
  return Excel.run(async (context) => {
    const styles = context.workbook.styles;
    styles.load("items/name");
    const registry = queueRegistryRead(context);
    const slot = await readActiveCell(context);
    const names = styles.items.map((style) => style.name);
    const checked = checkStyleName(rawName, names);
    if (!checked.ok) throw new Error(checked.message);
    // A style deleted in Excel's own Cell Styles leaves the list with this write.
    const entries = onlyExisting(registryEntries(registry), names);

    const spec = fromPaintSlot(checked.name, slot);
    styles.add(spec.name);
    // The list is written right behind the add: a batch the host stops halfway
    // through the properties leaves a style in Cell Styles, and it has to be one
    // the pane lists, so it can be updated or deleted instead of haunting it.
    queueRegistryWrite(
      context,
      withStyle(entries, { name: spec.name, kind: "user" }),
    );
    writeStyle(styles.getItem(spec.name), spec);
    await syncWrite(context, STAGE_SAVE);
    return `Saved style ${spec.name}`;
  });
}

/** Puts the style on every area of the selection, with pls,fix Undo. */
export async function applyCustomStyle(name: string): Promise<string> {
  return Excel.run(async (context) => {
    const found = probe(context, name);
    const registry = queueRegistryRead(context);
    // The areas' own sync carries the two reads above.
    const areas = await selectedAreas(context, STAGE_APPLY);
    if (found.isNullObject) {
      await forget(context, registry, name, STAGE_APPLY);
      throw new Error(gone(name));
    }
    await putStyleOnAreas(context, areas, found.name, STAGE_APPLY);
    return `Style: ${found.name}`;
  });
}

// How many cells wear the style, or null when a sheet was too large to read.
async function wearersOrNull(
  context: Excel.RequestContext,
  name: string,
  maxCells: number,
  maxTotalCells: number,
): Promise<number | null> {
  try {
    const { counts, skippedSheets } = await scanStyleWearers(
      context,
      maxCells,
      maxTotalCells,
    );
    return skippedSheets.length > 0
      ? null
      : (counts.get(name.toLowerCase()) ?? 0);
  } catch {
    return null;
  }
}

function updatedReceipt(name: string, worn: number | null): string {
  if (worn === null) return `Updated ${name}.`;
  if (worn === 0) return `Updated ${name}: no cell wears it yet.`;
  const cells = worn === 1 ? "1 cell wears" : `${countText(worn)} cells wear`;
  return `Updated ${name}: ${cells} it.`;
}

// A starter updated from a cell is the user's own from then on: the list files it
// as a user style, so Starter styles from palette leaves it alone. True when the
// write was queued; a list a newer pls,fix wrote is never written over.
function refileStarter(
  context: Excel.RequestContext,
  registry: Excel.Setting,
  name: string,
): boolean {
  const read = registryOf(registry);
  if (!read.ok) return false;
  const listed = read.entries.find((entry) => sameName(entry.name, name));
  if (listed?.kind !== "starter") return false;
  queueRegistryWrite(
    context,
    withStyle(read.entries, { name: listed.name, kind: "user" }),
  );
  return true;
}

const refiledReceipt = (name: string): string =>
  `Updated ${name}: it is yours now, Starter styles from palette leaves it alone.`;

/**
 * Rewrites the style from the active cell: every cell wearing it follows, which
 * is the point. The count of wearers is stated when the scan is cheap. A starter
 * style becomes the user's own, and the receipt says so instead of counting.
 */
export async function updateCustomStyle(
  name: string,
  maxCells = SHEET_SCAN_CELL_CAP,
  maxTotalCells = SHEET_SCAN_CELL_CAP,
): Promise<string> {
  return Excel.run(async (context) => {
    const found = probe(context, name);
    const registry = queueRegistryRead(context);
    const slot = await readActiveCell(context);
    if (found.isNullObject) {
      await forget(context, registry, name, STAGE_UPDATE);
      throw new Error(gone(name));
    }
    // The list is written first, as Save does: a batch the host stops halfway
    // must never leave a starter the next Starter styles press would repaint.
    const refiled = refileStarter(context, registry, found.name);
    writeStyle(found, fromPaintSlot(found.name, slot));
    await syncWrite(context, STAGE_UPDATE);
    if (refiled) return refiledReceipt(found.name);
    const worn = await wearersOrNull(
      context,
      found.name,
      maxCells,
      maxTotalCells,
    );
    return updatedReceipt(found.name, worn);
  });
}

// ---------------------------------------------------------------------------
// Delete
// ---------------------------------------------------------------------------

// Deleting a style sends every wearer back to Normal, so the cells are counted
// first, under the cap Find and the model check scan under (a deal model's big
// sheets fit, the scrubber's 5 000 would refuse most of them), and a workbook
// too large to read even then is refused outright: a style could be worn there.
async function requireWearers(
  context: Excel.RequestContext,
  name: string,
  maxCells: number,
  maxTotalCells: number,
): Promise<number> {
  const { counts, skippedSheets } = await scanStyleWearers(
    context,
    maxCells,
    maxTotalCells,
  );
  if (skippedSheets.length > 0) {
    throw new Error(
      `Some sheets are too large to scan (${skippedSheets.join(", ")}), so pls,fix cannot say how many cells wear the style. Delete it in Excel's Cell Styles instead.`,
    );
  }
  return counts.get(name.toLowerCase()) ?? 0;
}

/** How many cells a delete would send back to Normal; refuses a workbook too large to scan. */
export async function countStyleWearers(
  name: string,
  maxCells = SHEET_SCAN_CELL_CAP,
  maxTotalCells = SHEET_SCAN_CELL_CAP,
): Promise<number> {
  return Excel.run((context) =>
    requireWearers(context, name, maxCells, maxTotalCells),
  );
}

function deletedReceipt(name: string, worn: number): string {
  if (worn === 0) return `Deleted style ${name}.`;
  const cells = worn === 1 ? "1 cell went" : `${countText(worn)} cells went`;
  return `Deleted style ${name}: ${cells} back to Normal.`;
}

/** Deletes the style, counted first, and drops it from the list. */
export async function deleteCustomStyle(
  name: string,
  maxCells = SHEET_SCAN_CELL_CAP,
  maxTotalCells = SHEET_SCAN_CELL_CAP,
): Promise<string> {
  return Excel.run(async (context) => {
    const found = probe(context, name);
    const registry = queueRegistryRead(context);
    await context.sync();
    if (found.isNullObject) {
      await forget(context, registry, name, STAGE_DELETE);
      return `The style ${name} was already gone from this workbook.`;
    }
    if (found.builtIn) {
      throw new Error(
        `${found.name} is one of Excel's own styles and cannot be deleted.`,
      );
    }
    const worn = await requireWearers(
      context,
      found.name,
      maxCells,
      maxTotalCells,
    );
    const read = registryOf(registry);
    found.delete();
    if (read.ok)
      queueRegistryWrite(context, withoutStyle(read.entries, found.name));
    await syncWrite(context, STAGE_DELETE);
    return deletedReceipt(found.name, worn);
  });
}

// ---------------------------------------------------------------------------
// Starter styles
// ---------------------------------------------------------------------------

// The shared wording first; then the styles the list files as the user's (an updated starter),
// then the ones the list does not hold at all.
function starterReceipt(
  added: number,
  restyled: number,
  yours: string[],
  notMade: string[],
): string {
  const alone = yours.length === 0 ? "" : yoursNow(yours);
  const others =
    notMade.length === 0
      ? ""
      : ` Left alone ${notMade.join(", ")}: already in this workbook, not made by pls,fix.`;
  return startersDone(added, restyled) + alone + others;
}

/**
 * The eight starter styles from the CURRENT palette: the missing ones are
 * created and the ones pls,fix made are rewritten, so pressing it after a
 * palette change restyles every cell wearing them. Idempotent. A style of a
 * starter's name is left exactly as it is, and the receipt names it: one the
 * list files as the user's (a starter they updated) is "yours now", one the
 * list does not hold (a colleague's, from Excel's own gallery) was "not made
 * by pls,fix".
 */
export async function addStarterStyles(): Promise<string> {
  return Excel.run(async (context) => {
    const styles = context.workbook.styles;
    styles.load("items/name");
    const registry = queueRegistryRead(context);
    await context.sync();
    const names = styles.items.map((style) => style.name);
    let entries = onlyExisting(registryEntries(registry), names);

    let added = 0;
    let restyled = 0;
    const yours: string[] = [];
    const notMade: string[] = [];
    for (const spec of starterStyles(getActiveSettings())) {
      const exists = names.some((name) => sameName(name, spec.name));
      const listed = entries.find((entry) => sameName(entry.name, spec.name));
      if (exists && listed?.kind !== "starter") {
        if (listed === undefined) notMade.push(spec.name);
        else yours.push(listed.name);
        continue;
      }
      if (exists) restyled += 1;
      else {
        styles.add(spec.name);
        added += 1;
      }
      writeStyle(styles.getItem(spec.name), spec);
      entries = withStyle(entries, { name: spec.name, kind: "starter" });
    }
    queueRegistryWrite(context, entries);
    await syncWrite(context, STAGE_STARTERS);
    return starterReceipt(added, restyled, yours, notMade);
  });
}

// ---------------------------------------------------------------------------
// The list
// ---------------------------------------------------------------------------

export interface CustomStyleList {
  styles: StylePreview[];
  /** Why the list is empty when it is not because there are no styles. */
  note?: string;
}

function previewOf(entry: RegistryEntry, style: Excel.Style): StylePreview {
  const fill = style.includePatterns
    ? readFill({ color: style.fill.color, pattern: style.fill.pattern })
    : null;
  return {
    name: style.name,
    kind: entry.kind,
    fill: fill?.color ? fill.color.toUpperCase() : null,
    fontColor: style.includeFont ? style.font.color.toUpperCase() : null,
    bold: style.includeFont && style.font.bold,
  };
}

/**
 * The styles pls,fix made that the workbook still has, in list order, each with
 * what a row needs to draw its preview. Read-only: a style gone from Excel is
 * left out, and the list is tidied the next time something writes it.
 */
export async function listCustomStyles(): Promise<CustomStyleList> {
  return Excel.run(async (context) => {
    const styles = context.workbook.styles;
    styles.load("items/name");
    const registry = queueRegistryRead(context);
    await context.sync();
    const read = registryOf(registry);
    if (!read.ok) return { styles: [], note: read.message };

    const shown = onlyExisting(
      read.entries,
      styles.items.map((style) => style.name),
    ).map((entry) => {
      const style = styles.getItem(entry.name);
      style.load(
        "name,includeFont,includePatterns,fill/color,fill/pattern,font/color,font/bold",
      );
      return { entry, style };
    });
    await context.sync();
    return {
      styles: shown.map(({ entry, style }) => previewOf(entry, style)),
    };
  });
}
