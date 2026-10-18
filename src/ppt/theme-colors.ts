// The brand colors in the deck's own theme (PowerPointApi 1.10): Accent 1 to 6 of
// every slide master are read and saved in the deck's design part FIRST, then
// overwritten with the brand, so PowerPoint's own color picker and every object
// drawn in theme colors (the tables' Medium Style 2 Accent 1) follow; Restore
// writes the saved colors back and clears them. Owns every ThemeColorScheme call.
// Invariants: the colors are saved before they are touched and cleared only after
// they are back; Accent 1 to 6 only; a refused batch is not atomic, so after a
// refusal the saved colors are written back in a run of their own.

import { normalizeHex } from "../settings";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import {
  adoptDeckDesign,
  deckDesign,
  distrustDeckDesign,
  loadDeckDesign,
  saveDeckDesign,
  settleDeckDesign,
  writeDeckDesign,
  type DesignWrite,
} from "./deck-design";
import {
  buildDeckDesign,
  THEME_ACCENTS,
  type DeckDesign,
  type SavedTheme,
} from "./deck-design-record";
import { nudgeSave } from "./save-nudge";
import { hasPowerPointApi } from "./shapes";

export const THEME_API = "1.10";
export const NEEDS_THEME_API =
  "The deck theme needs Microsoft 365 2601 / Mac 16.105 or newer.";
export const THEME_REFUSED =
  "PowerPoint refused the theme colors; the deck keeps its own.";
// A put the host refused halfway and a write-back it refused too: the theme may
// show some brand colors, so Restore stays and the sentence does not say "its own".
export const THEME_PARTLY =
  "PowerPoint refused some of the theme colors, so some accents may now hold brand colors. Restore theme colors puts the saved ones back.";
export const THEME_RESTORED = "Theme colors restored.";
export const NOTHING_SAVED =
  "Nothing to restore: pls,fix has not changed this deck's theme colors.";
export const NO_SAVED_MASTER =
  "The slide masters those colors were saved from are no longer in this deck.";

// Accent 1 to 6, in the order the brand fills them: brand[0] is Accent 1.
const ACCENT_NAMES = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
] as const;

const READ_MASTERS = "reading the slide masters";
const READ_COLORS = "reading the theme colors";
const PUT_COLORS = "putting the brand colors in the theme";
const BACK_COLORS = "putting the theme colors back";
const RESTORE_COLORS = "restoring the theme colors";

export function hasThemeApi(): boolean {
  return hasPowerPointApi(THEME_API);
}

function requireThemeApi(): void {
  if (!hasThemeApi()) throw new Error(NEEDS_THEME_API);
}

/** True when the deck's design part holds the colors a put replaced. */
export function hasSavedTheme(): boolean {
  return (deckDesign()?.themes.length ?? 0) > 0;
}

/** The receipt of a put: how many accents the brand filled. */
export function themeReceipt(slots: number): string {
  return `Brand colors are in this deck's theme: Accent 1-${String(slots)}.`;
}

// ---------------------------------------------------------------------------
// Reading
// ---------------------------------------------------------------------------

// A color the host gave back, "#RRGGBB" or the bare "RRGGBB" its docs show.
function asColor(value: string): string {
  const hex = normalizeHex(value);
  if (hex === null) {
    throw new Error("PowerPoint gave back a theme color that is not a color.");
  }
  return hex;
}

/** Accent 1 to 6 of the first slide master: one round trip. */
export async function readThemeAccents(): Promise<string[]> {
  requireThemeApi();
  return PowerPoint.run(async (context) => {
    const scheme =
      context.presentation.slideMasters.getItemAt(0).themeColorScheme;
    const reads = ACCENT_NAMES.map((name) => scheme.getThemeColor(name));
    await withSyncDeadline(context.sync(), READ_COLORS);
    return reads.map((read) => asColor(read.value));
  });
}

// The six accents of every master, each under its master's id: two round trips.
async function readEveryMaster(): Promise<SavedTheme[]> {
  return PowerPoint.run(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await withSyncDeadline(context.sync(), READ_MASTERS);
    const reads = masters.items.map((master) => ({
      master: master.id,
      colors: ACCENT_NAMES.map((name) =>
        master.themeColorScheme.getThemeColor(name),
      ),
    }));
    await withSyncDeadline(context.sync(), READ_COLORS);
    return reads.map(({ master, colors }) => ({
      master,
      accents: colors.map((color) => asColor(color.value)),
    }));
  });
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

function writeAccents(
  context: PowerPoint.RequestContext,
  master: string,
  colors: readonly string[],
): void {
  const scheme =
    context.presentation.slideMasters.getItem(master).themeColorScheme;
  colors.forEach((color, at) => {
    scheme.setThemeColor(ACCENT_NAMES[at]!, color);
  });
}

// What a master had before pls,fix first touched it: a scheme an earlier put
// saved stays the original, so putting again never saves the brand over it.
function keepOriginals(
  saved: readonly SavedTheme[],
  current: readonly SavedTheme[],
): SavedTheme[] {
  return current.map(
    (now) => saved.find((was) => was.master === now.master) ?? now,
  );
}

// A refusal is not atomic: whatever the batch did before it stays done. One batch
// of its own writes the colors back from what was read, then the part as it was.
// True when the part's write landed with it.
async function backBatch(
  before: readonly SavedTheme[],
  restored: DeckDesign | null,
): Promise<boolean> {
  try {
    const write = await PowerPoint.run(async (context) => {
      for (const { master, accents } of before) {
        writeAccents(context, master, accents);
      }
      const queued =
        restored === null ? null : writeDeckDesign(context, restored);
      await withSyncDeadline(context.sync(), BACK_COLORS);
      return queued !== null && queued.ok ? queued.write : null;
    });
    if (write === null) return false;
    adoptDeckDesign(write);
    return true;
  } catch {
    distrustDeckDesign(); // the batch may have written the part before it refused
    return false;
  }
}

// What the theme shows now, null when the host will not say.
async function themeNow(): Promise<SavedTheme[] | null> {
  try {
    return await readEveryMaster();
  } catch {
    return null;
  }
}

// True when every master the deck still has shows the colors it had before.
function showsOwn(
  shown: readonly SavedTheme[],
  before: readonly SavedTheme[],
): boolean {
  return before.every((was) => {
    const now = shown.find((one) => one.master === was.master);
    return (
      now === undefined ||
      was.accents.every((color, at) => color === now.accents[at])
    );
  });
}

// After a refused put: the colors are written back, then the theme is read to see
// what it shows. On its own colors again the put's saved colors are dropped from
// the part, so Restore is not left offering to undo nothing; anything else, or a
// theme that cannot be read, keeps them for Restore. True = the deck is as it was.
async function writeBack(
  before: readonly SavedTheme[],
  restored: DeckDesign | null,
): Promise<boolean> {
  distrustDeckDesign();
  await loadDeckDesign();
  const landed = await backBatch(before, restored);
  const shown = await themeNow();
  const own = shown !== null && showsOwn(shown, before);
  if (own && !landed && restored !== null) {
    await saveDeckDesign(() => restored);
  } else if (!landed) {
    await loadDeckDesign();
  }
  return own;
}

// One batch: the design part (saved first), then every master's accents. A note
// instead of a write when the part cannot be written: no theme write was queued.
async function putBatch(
  design: DeckDesign,
  masters: readonly SavedTheme[],
  colors: readonly string[],
): Promise<DesignWrite | string> {
  return PowerPoint.run(async (context) => {
    const queued = writeDeckDesign(context, design);
    if (!queued.ok) return queued.note;
    for (const { master } of masters) writeAccents(context, master, colors);
    await withSyncDeadline(context.sync(), PUT_COLORS);
    return queued.write;
  });
}

async function readBeforeTouching(): Promise<SavedTheme[]> {
  try {
    return await readEveryMaster();
  } catch (error) {
    if (isDrawTimeout(error)) throw error;
    throw new Error(THEME_REFUSED);
  }
}

// A put and a restore build on the part the deck holds, so a copy that is in
// doubt (never read, or its read refused) is read again first; a deck that still
// cannot be read is left exactly as it is.
async function requireDeckRead(): Promise<void> {
  const unread = await settleDeckDesign();
  if (unread !== null) throw new Error(`${unread} Nothing was changed.`);
}

/**
 * Writes the brand into Accent 1 to min(6, count) of every slide master, after
 * saving what each had in the deck's design part. Answers how many accents the
 * brand filled. Refused without a round trip below PowerPointApi 1.10; a read or
 * a write the host refuses leaves the deck with its own colors.
 */
export async function putBrandInTheme(
  brand: readonly string[],
): Promise<number> {
  requireThemeApi();
  await requireDeckRead();
  const before = await readBeforeTouching();
  const previous = deckDesign()?.themes ?? [];
  const design = buildDeckDesign(brand, keepOriginals(previous, before));
  if (design === null) throw new Error(THEME_REFUSED);
  const slots = Math.min(THEME_ACCENTS, design.brand.length);
  let done: DesignWrite | string;
  try {
    done = await putBatch(design, before, design.brand.slice(0, slots));
  } catch (error) {
    distrustDeckDesign();
    if (isDrawTimeout(error)) throw error;
    const back = await writeBack(
      before,
      buildDeckDesign(design.brand, previous),
    );
    throw new Error(back ? THEME_REFUSED : THEME_PARTLY);
  }
  if (typeof done === "string") throw new Error(done);
  adoptDeckDesign(done);
  await nudgeSave();
  return slots;
}

// ---------------------------------------------------------------------------
// Restoring
// ---------------------------------------------------------------------------

// One batch: the saved colors of every master still in the deck, then the part
// without them - after them, so a refused write leaves the saved scheme where it
// was. How many masters were written, and the part's write.
async function restoreBatch(
  saved: readonly SavedTheme[],
  cleared: DeckDesign,
): Promise<{ written: number; write: DesignWrite | string }> {
  return PowerPoint.run(async (context) => {
    const masters = context.presentation.slideMasters;
    masters.load("items/id");
    await withSyncDeadline(context.sync(), READ_MASTERS);
    const here = new Set(masters.items.map((master) => master.id));
    const present = saved.filter((theme) => here.has(theme.master));
    for (const { master, accents } of present) {
      writeAccents(context, master, accents);
    }
    const queued = writeDeckDesign(context, cleared);
    await withSyncDeadline(context.sync(), RESTORE_COLORS);
    return {
      written: present.length,
      write: queued.ok ? queued.write : queued.note,
    };
  });
}

/**
 * Writes the colors the deck's theme had back, every master that is still
 * there, and clears the saved scheme. Answers the receipt. Nothing to restore
 * says so; a refusal keeps the saved scheme so it can be pressed again.
 */
export async function restoreThemeColors(): Promise<string> {
  requireThemeApi();
  await requireDeckRead();
  const design = deckDesign();
  if (design === null || design.themes.length === 0) {
    throw new Error(NOTHING_SAVED);
  }
  const cleared = buildDeckDesign(design.brand, []);
  if (cleared === null) throw new Error(THEME_REFUSED);
  let outcome: Awaited<ReturnType<typeof restoreBatch>>;
  try {
    outcome = await restoreBatch(design.themes, cleared);
  } catch (error) {
    distrustDeckDesign();
    if (isDrawTimeout(error)) throw error;
    await loadDeckDesign();
    throw new Error(THEME_REFUSED);
  }
  if (typeof outcome.write === "string") {
    distrustDeckDesign();
    await loadDeckDesign();
    throw new Error(outcome.write);
  }
  adoptDeckDesign(outcome.write);
  await nudgeSave();
  return outcome.written === 0 ? NO_SAVED_MASTER : THEME_RESTORED;
}
