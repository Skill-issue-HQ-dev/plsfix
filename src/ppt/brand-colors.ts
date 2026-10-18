// The PowerPoint pane's brand colors: the Tools tab's one-click swatches (Fill,
// Outline, Text), the Settings tab's palette editor with its four buttons, and
// the palette that arrives with a link paste. Owns the wiring and the palette in
// memory; the rules are brand-palette.ts, the shapes shape-paint.ts, the deck's
// part deck-design.ts, the theme theme-colors.ts and theme-palette.ts. Takes the
// pane's `act` the way chooser.ts does, so a press runs under its guard and its
// busy latch. Invariants: the palette shown is drawn from the shipped two before
// any host is read; every flow that writes the deck's part runs inside its lock.

import type { KeyStore } from "../link/workspace";
import { brandColors, withBrandList } from "../palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "../settings";
import { getElement } from "../ui/dom";
import { renderPaletteEditor } from "../ui/palette-editor";
import { renderSwatches, type SwatchStrip } from "../ui/swatches";
import type { Toast } from "../ui/toast";
import {
  chooseBrand,
  paletteFromAccents,
  readUserDefault,
  samePalette,
} from "./brand-palette";
import { saveBrand } from "./brand-save";
import {
  deckDesign,
  deckDesignUnknown,
  designAvailable,
  loadDeckDesign,
  settleDeckDesign,
  unreadNote,
  withDesignLock,
} from "./deck-design";
import { paintShapes } from "./shape-paint";
import {
  hasSavedTheme,
  hasThemeApi,
  NEEDS_THEME_API,
  putBrandInTheme,
  restoreThemeColors,
  themeReceipt,
} from "./theme-colors";
import { themeAccents, themeBrand } from "./theme-palette";

// Fill and Outline can be cleared, a text color cannot: the host has no
// "automatic" to write.
export const PPT_SWATCH_STRIPS: readonly SwatchStrip[] = [
  { mode: "fill", label: "Fill", none: "No fill" },
  { mode: "outline", label: "Outline", none: "No outline" },
  { mode: "text", label: "Text" },
];

const THEME_HINT =
  "Put in the deck theme fills Accent 1 to 6, so PowerPoint's own color picker offers your brand. Restore puts the old colors back.";
const ONE_COLOR_THEME =
  "This deck's theme holds fewer than two different accent colors.";
const PASTE_FIRST = "Paste links from Excel first.";

export interface BrandColorsDeps {
  /** The pane's guarded runner: busy state, toast, error reporting. */
  act(run: () => Promise<string>, action: string): void;
  /** True while a guarded action runs: what the buttons it made are born as. */
  busy(): boolean;
  /** A line for the toast's details, beside the receipt a flow answers. */
  note(line: string): void;
  /** For what no guarded action is there to say: a save that failed. */
  toast: Toast;
  keyStore: KeyStore;
}

export interface BrandColors {
  /** The boot read: the deck's palette, the user's default, the deck's theme. */
  boot(): Promise<void>;
  /**
   * A copy's brand colors, after a link paste. Answers the sentence the paste's
   * receipt takes on ("" when there is nothing to say).
   */
  receive(palette: readonly string[] | undefined): Promise<string>;
  /** Draws the swatches and the editor from this list, saving nothing. */
  show(brand: readonly string[]): void;
  /** The palette the pane shows now, with the modelling colors under it. */
  settings(): BrandSettings;
}

interface State {
  deps: BrandColorsDeps;
  /** The palette shown, with the user default's other settings under it. */
  settings: BrandSettings;
  /** The user's default as last read or written: the palette of a deck that cannot keep one. */
  userBrand: string[] | null;
  /** The last palette a link paste carried. */
  workbook: string[] | null;
  /** Whether the host has PowerPointApi 1.10: unknown (null) until the boot asks. */
  themeApi: boolean | null;
  /** True once the boot has read the deck: before it, an unread deck is just "not yet". */
  booted: boolean;
}

const button = (id: string): HTMLButtonElement =>
  getElement<HTMLButtonElement>(id);

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// ---------------------------------------------------------------------------
// Drawing
// ---------------------------------------------------------------------------

function drawSwatches(state: State): void {
  renderSwatches(getElement("ppt-brand-swatches"), {
    colors: brandColors(state.settings),
    strips: PPT_SWATCH_STRIPS,
    onPick: (mode, hex) => {
      state.deps.act(() => paintShapes(mode, hex), `swatch-${mode}`);
    },
  });
}

// A control made while an action runs has to come back as disabled as the ones
// it replaced: setBusy has never seen it.
function keepBusy(state: State): void {
  if (!state.deps.busy()) return;
  for (const each of document.querySelectorAll<HTMLButtonElement>(
    "#ppt-brand-swatches button, #ppt-brand-editor button",
  )) {
    each.disabled = true;
  }
}

function workbookDiffers(state: State): boolean {
  return (
    state.workbook !== null &&
    !samePalette(state.workbook, state.settings.brand)
  );
}

// setBusy enables every button when an action ends, so a button that is dead for
// a reason of its own says so with data-dead, which it leaves alone.
function setLive(id: string, live: boolean, busy: boolean): void {
  const node = button(id);
  node.toggleAttribute("data-dead", !live);
  node.disabled = !live || busy;
}

// Restore stays pressable while the deck's saved colors are unknown: a press
// reads the deck again, and says so when it still cannot.
function restorable(state: State): boolean {
  return hasSavedTheme() || (state.booted && deckDesignUnknown());
}

function syncButtons(state: State): void {
  const busy = state.deps.busy();
  const theme = state.themeApi !== false;
  setLive("use-workbook-colors", workbookDiffers(state), busy);
  setLive("use-deck-theme-colors", theme, busy);
  setLive("put-theme-colors", theme, busy);
  setLive("restore-theme-colors", theme && restorable(state), busy);
  getElement("ppt-theme-hint").textContent = theme
    ? THEME_HINT
    : NEEDS_THEME_API;
}

function render(state: State): void {
  drawSwatches(state);
  renderPaletteEditor(getElement("ppt-brand-editor"), {
    settings: state.settings,
    firstRow: 0,
    onPreview: (next) => {
      state.settings = next;
      drawSwatches(state);
      keepBusy(state);
    },
    onCommit: (next) => commit(state, next),
  });
  keepBusy(state);
  syncButtons(state);
}

// ---------------------------------------------------------------------------
// Saving: brand-save.ts writes the deck's part and the user's default
// ---------------------------------------------------------------------------

// One finished edit: the notes are what did not land.
async function persist(
  state: State,
  brand: readonly string[],
): Promise<string[]> {
  const saved = await saveBrand(state.deps.keyStore, brand);
  if (saved.user !== null) state.userBrand = saved.user;
  return saved.notes;
}

function say(state: State, notes: readonly string[]): void {
  const [first, ...rest] = notes;
  if (first === undefined) return;
  state.deps.toast.show(first, "error", rest.join("\n") || undefined);
}

function commit(state: State, next: BrandSettings): void {
  state.settings = next;
  render(state);
  void persist(state, next.brand).then(
    (notes) => {
      say(state, notes);
    },
    (error: unknown) => {
      say(state, [reasonOf(error)]);
    },
  );
}

async function adopt(state: State, palette: readonly string[]): Promise<void> {
  const next = withBrandList(state.settings, palette);
  if (next === null) throw new Error("A palette needs at least two colors.");
  state.settings = next;
  render(state);
  for (const note of await persist(state, next.brand)) state.deps.note(note);
}

// ---------------------------------------------------------------------------
// The four buttons
// ---------------------------------------------------------------------------

async function useWorkbook(state: State): Promise<string> {
  const palette = state.workbook;
  if (palette === null) throw new Error(PASTE_FIRST);
  await adopt(state, palette);
  return `Brand colors from the workbook: ${String(palette.length)}.`;
}

async function useDeckTheme(state: State): Promise<string> {
  const palette = paletteFromAccents(await themeAccents());
  if (palette === null) throw new Error(ONE_COLOR_THEME);
  await adopt(state, palette);
  return `Brand colors from the deck's theme: ${String(palette.length)}.`;
}

async function putTheme(state: State): Promise<string> {
  try {
    const brand = state.settings.brand;
    return themeReceipt(await withDesignLock(() => putBrandInTheme(brand)));
  } finally {
    syncButtons(state);
  }
}

async function restoreTheme(state: State): Promise<string> {
  try {
    return await withDesignLock(restoreThemeColors);
  } finally {
    syncButtons(state);
  }
}

// Edit colors: the Settings tab, brought to the palette, the first picker in focus.
function openPalette(): void {
  button("tab-settings").click();
  getElement("ppt-brand-heading").scrollIntoView?.({ block: "start" });
  document
    .querySelector<HTMLElement>('#ppt-brand-editor input[type="color"]')
    ?.focus();
}

function wire(state: State): void {
  button("edit-brand-colors").addEventListener("click", openPalette);
  const flows: [string, (state: State) => Promise<string>][] = [
    ["use-workbook-colors", useWorkbook],
    ["use-deck-theme-colors", useDeckTheme],
    ["put-theme-colors", putTheme],
    ["restore-theme-colors", restoreTheme],
  ];
  for (const [id, run] of flows) {
    button(id).addEventListener("click", () => {
      state.deps.act(() => run(state), id);
    });
  }
}

// ---------------------------------------------------------------------------
// The boot and the paste
// ---------------------------------------------------------------------------

async function boot(state: State): Promise<void> {
  state.themeApi = hasThemeApi();
  const user = await readUserDefault(state.deps.keyStore);
  const read = await withDesignLock(loadDeckDesign);
  if (read.issue !== undefined) state.deps.toast.show(read.issue, "error");
  if (read.failure !== undefined) {
    state.deps.toast.show(unreadNote(read.failure), "error");
  }
  // An unread deck may hold a palette: its theme's accents are not offered.
  const needsTheme =
    read.failure === undefined && read.design === null && user === null;
  const brand = chooseBrand({
    deck: read.design?.brand,
    user: user?.brand,
    theme: needsTheme ? await themeBrand(state.themeApi) : null,
  });
  state.settings =
    withBrandList(user ?? DEFAULT_SETTINGS, brand) ?? state.settings;
  state.userBrand = user === null ? null : [...user.brand];
  state.booted = true;
  render(state);
}

const UNREAD = "unread";
const PASTE_UNREAD =
  " The workbook's brand colors were not applied, because this deck's own could not be read. Use the workbook's colors takes them.";

// The deck's own palette: its part, or - on a host that cannot keep one - the
// user's default. A deck that could not be read is not "none": it is read once
// more, and when that fails too the answer is UNREAD.
async function ownPalette(
  state: State,
): Promise<readonly string[] | null | typeof UNREAD> {
  if (!designAvailable()) return state.userBrand;
  const unread = await withDesignLock(settleDeckDesign);
  if (unread !== null) return UNREAD;
  return deckDesign()?.brand ?? null;
}

async function receive(
  state: State,
  palette: readonly string[] | undefined,
): Promise<string> {
  if (palette === undefined) return "";
  state.workbook = [...palette];
  const own = await ownPalette(state);
  if (own === UNREAD) {
    syncButtons(state);
    return PASTE_UNREAD;
  }
  if (own === null) {
    await adopt(state, palette);
    return ` Brand colors from the workbook: ${String(palette.length)}.`;
  }
  syncButtons(state);
  if (samePalette(own, palette)) return "";
  return " The workbook's brand colors differ from this deck's.";
}

function showBrand(state: State, brand: readonly string[]): void {
  state.settings = withBrandList(state.settings, brand) ?? state.settings;
  render(state);
}

let installed: State | null = null;

/** Draws the swatches and the editor from a list, saving nothing (a gate's seed). */
export function drawBrandColors(brand: readonly string[]): void {
  if (installed !== null) showBrand(installed, brand);
}

/**
 * Wires the Tools tab's swatches, Edit colors and the Settings section's four
 * buttons, and draws the shipped palette at once: no host has been read yet.
 */
export function installBrandColors(deps: BrandColorsDeps): BrandColors {
  const state: State = {
    deps,
    settings: { ...DEFAULT_SETTINGS },
    userBrand: null,
    workbook: null,
    themeApi: null,
    booted: false,
  };
  installed = state;
  wire(state);
  render(state);
  return {
    boot: () => boot(state),
    receive: (palette) => receive(state, palette),
    show: (brand) => {
      showBrand(state, brand);
    },
    settings: () => state.settings,
  };
}
