// The palette editor both panes share: brand rows past Primary and Accent (every
// row where a pane has no pickers for those two), Add color, Paste colors. DOM
// only: onPreview and onCommit go to the host. Invariant: one listener set per
// container; a typed-in control is kept; a refused row edit gives its color back.

import {
  brandColors,
  heldElsewhere,
  MAX_BRAND_COLORS,
  parseColorList,
  repeatSentence,
  shadeLadder,
  withBrandColor,
  withBrandList,
  withoutBrandColor,
} from "../palette";
import { normalizeHex, type BrandSettings } from "../settings";

export interface PaletteEditorOptions {
  /** The active settings: the rows show brand colors 3 and up. */
  settings: BrandSettings;
  /**
   * The first brand color that gets a row: 2 (the default) leaves Primary and
   * Accent to the pane's own pickers, 0 shows every color. Fixed per container.
   */
  firstRow?: number;
  /** The settings as they stand mid-edit (a picker being dragged). */
  onPreview: (settings: BrandSettings) => void;
  /** The settings once an edit is finished: save them. */
  onCommit: (settings: BrandSettings) => void;
}

interface State {
  options: PaletteEditorOptions;
  settings: BrandSettings;
  rows: HTMLElement;
  paste: HTMLInputElement;
  status: HTMLElement;
  /**
   * The list as it stood before the first preview of an edit that no commit has
   * saved yet (null while none is out): what a refused edit puts back.
   */
  before: BrandSettings | null;
}

const FIRST_EXTRA = 2;
const SEED_COLOR = "#808080";
const MAX_WORD = "Ten colors is the most";
const states = new WeakMap<HTMLElement, State>();

function make<K extends keyof HTMLElementTagNameMap>(
  doc: Document,
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = doc.createElement(tag);
  if (className !== undefined) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function field(
  doc: Document,
  className: string,
  label: string,
): HTMLInputElement {
  const input = make(doc, "input", className);
  input.setAttribute("aria-label", label);
  return input;
}

function removeButton(doc: Document, name: string): HTMLButtonElement {
  const remove = make(doc, "button", "icon-button", "×");
  remove.type = "button";
  remove.setAttribute("aria-label", `Remove ${name}`);
  remove.title = `Remove ${name}`;
  return remove;
}

function rowNode(doc: Document, index: number): HTMLElement {
  const name = `Brand ${String(index + 1)}`;
  const row = make(doc, "div", "palette-row extra");
  row.dataset.brandIndex = String(index);
  const picker = field(doc, "swatch-input", `${name} color picker`);
  picker.type = "color";
  const meta = make(doc, "span", "palette-meta");
  meta.append(make(doc, "strong", undefined, name));
  const hex = field(doc, "hex-input", `${name} hex value`);
  hex.type = "text";
  hex.maxLength = 7;
  hex.spellcheck = false;
  row.append(picker, meta, hex);
  // Primary and Accent are what a palette always has: nothing to remove.
  if (index >= FIRST_EXTRA) row.append(removeButton(doc, name));
  return row;
}

// Rows follow the settings in place: a control the user holds is left alone,
// unless `force` (a refused edit gives the row its color back, held or not).
function fillRow(row: HTMLElement, hex: string, force = false): void {
  const held = row.ownerDocument.activeElement;
  const picker = row.querySelector<HTMLInputElement>('input[type="color"]')!;
  const text = row.querySelector<HTMLInputElement>('input[type="text"]')!;
  if (force || picker !== held) picker.value = hex.toLowerCase();
  if (force || text !== held) text.value = hex;
}

function syncRows(state: State): void {
  const doc = state.rows.ownerDocument;
  const first = Math.max(0, state.options.firstRow ?? FIRST_EXTRA);
  const shown = brandColors(state.settings).slice(first);
  const rows = [...state.rows.children] as HTMLElement[];
  shown.forEach((hex, at) => {
    let row = rows[at];
    if (!row) {
      row = rowNode(doc, at + first);
      state.rows.append(row);
    }
    fillRow(row, hex);
  });
  for (const stale of rows.slice(shown.length)) stale.remove();
}

function say(state: State, message: string): string {
  state.status.textContent = message;
  return message;
}

function sameList(first: BrandSettings, second: BrandSettings): boolean {
  const a = brandColors(first);
  const b = brandColors(second);
  return a.length === b.length && a.every((color, at) => color === b[at]);
}

function preview(state: State, next: BrandSettings | null): void {
  if (!next || sameList(next, state.settings)) return;
  state.before ??= state.settings;
  state.settings = next;
  state.options.onPreview(next);
}

function commit(state: State, next: BrandSettings | null): void {
  if (!next) return;
  // Saved once per edit: Enter then change, or a change with nothing new, adds none.
  if (state.before === null && sameList(next, state.settings)) return;
  state.before = null;
  state.settings = next;
  state.options.onCommit(next);
}

// A refused edit takes back what it previewed, and the host is told, so a strip
// never keeps a half-typed color.
function putBack(state: State): void {
  const back = state.before;
  if (back === null) return;
  state.before = null;
  state.settings = back;
  state.options.onPreview(back);
}

// A color not on the list yet to start a new row with: the lighter shade of
// the last one, the neutral grey if every shade is taken.
function newColor(settings: BrandSettings): string {
  const colors = brandColors(settings);
  const last = colors.at(-1) ?? SEED_COLOR;
  const shades = shadeLadder(last).map((shade) => shade.hex);
  return (
    [...shades, SEED_COLOR].find((hex) => !colors.includes(hex)) ?? SEED_COLOR
  );
}

function addColor(state: State): void {
  const next = withBrandColor(
    state.settings,
    brandColors(state.settings).length,
    newColor(state.settings),
  );
  if (!next) {
    say(state, `${MAX_WORD}, so none was added. Remove one first.`);
    return;
  }
  commit(state, next);
  say(state, "Added a color at the end of the list.");
}

// The one rule for colors that arrive in bulk (a paste, a read off the sheet):
// two or more replace the whole list, primary and accent included, one is
// added to the end, none says so. Answers the sentence it put on the status line.
function applyColors(
  state: State,
  colors: readonly string[],
  dropped: number,
  none: string,
): string {
  const [only] = colors;
  if (only === undefined) return say(state, none);
  if (colors.length === 1) return appendOne(state, only);
  commit(state, withBrandList(state.settings, colors));
  const more = dropped > 0 ? `, ${String(dropped)} more did not fit` : "";
  return say(state, `${String(colors.length)} colors read${more}.`);
}

function applyText(state: State, text: string): void {
  const { colors, dropped } = parseColorList(text);
  applyColors(state, colors, dropped, "No colors found in that text.");
}

function appendOne(state: State, hex: string): string {
  const colors = brandColors(state.settings);
  if (colors.includes(hex)) return say(state, repeatSentence(hex));
  const next = withBrandColor(state.settings, colors.length, hex);
  if (!next) {
    return say(
      state,
      `${MAX_WORD}, so ${hex} was not added. Remove one first.`,
    );
  }
  commit(state, next);
  return say(state, `Added ${hex}.`);
}

function rowOf(target: EventTarget | null): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>("[data-brand-index]")
    : null;
}

function onRowInput(state: State, event: Event): void {
  const row = rowOf(event.target);
  const input = event.target;
  if (!row || !(input instanceof HTMLInputElement)) return;
  const hex = normalizeHex(input.value);
  if (!hex) return;
  const index = Number(row.dataset.brandIndex);
  if (input.type === "text") {
    const picker = row.querySelector<HTMLInputElement>('input[type="color"]')!;
    picker.value = hex.toLowerCase();
  } else {
    row.querySelector<HTMLInputElement>('input[type="text"]')!.value = hex;
  }
  // A color the list holds elsewhere is refused when the edit finishes.
  if (heldElsewhere(state.settings, index, hex)) return;
  preview(state, withBrandColor(state.settings, index, hex));
}

// A refused row edit says why and gives the row the color it had.
function refuse(state: State, row: HTMLElement, at: number, why: string): void {
  putBack(state);
  say(state, why);
  const color = brandColors(state.settings)[at];
  if (color !== undefined) fillRow(row, color, true);
}

function finishRowEdit(
  state: State,
  row: HTMLElement,
  input: HTMLInputElement,
): void {
  const index = Number(row.dataset.brandIndex);
  const hex = normalizeHex(input.value);
  const text = row.querySelector<HTMLInputElement>('input[type="text"]')!;
  if (!hex) {
    refuse(
      state,
      row,
      index,
      `"${input.value}" is not a hex color like #6D7E12.`,
    );
    return;
  }
  if (heldElsewhere(state.settings, index, hex)) {
    refuse(state, row, index, repeatSentence(hex));
    return;
  }
  text.value = hex;
  commit(state, withBrandColor(state.settings, index, hex));
}

function onRowChange(state: State, event: Event): void {
  const row = rowOf(event.target);
  if (row && event.target instanceof HTMLInputElement) {
    finishRowEdit(state, row, event.target);
  }
}

function onRowKeyDown(state: State, event: KeyboardEvent): void {
  const row = rowOf(event.target);
  const input = event.target;
  if (!row || event.key !== "Enter" || !(input instanceof HTMLInputElement)) {
    return;
  }
  if (input.type === "text") finishRowEdit(state, row, input);
}

function onRowClick(state: State, event: Event): void {
  const row = rowOf(event.target);
  const remove = (event.target as Element).closest("button");
  if (!row || !remove) return;
  const index = Number(row.dataset.brandIndex);
  const gone = brandColors(state.settings)[index];
  commit(state, withoutBrandColor(state.settings, index));
  say(state, `Removed ${gone ?? "the color"}.`);
}

function takeText(state: State): void {
  const text = state.paste.value;
  state.paste.value = "";
  if (text.trim() !== "") applyText(state, text);
}

function build(container: HTMLElement, options: PaletteEditorOptions): State {
  const doc = container.ownerDocument;
  const rows = make(doc, "div", "palette-extra-rows");
  const add = make(doc, "button", "palette-add", "Add color");
  add.id = "add-brand-color";
  add.type = "button";
  const label = make(doc, "label", "key-label");
  const paste = field(doc, "key-input", "Paste colors");
  paste.id = "paste-brand-colors";
  paste.type = "text";
  paste.spellcheck = false;
  paste.placeholder = "Hex, rgb() or palette JSON";
  label.append(make(doc, "span", undefined, "Paste colors"), paste);
  const status = make(doc, "p", "hint");
  status.setAttribute("aria-live", "polite");
  status.setAttribute("role", "status");
  status.dataset.paletteStatus = "";
  const state: State = {
    options,
    settings: options.settings,
    rows,
    paste,
    status,
    before: null,
  };

  rows.addEventListener("input", (event) => onRowInput(state, event));
  rows.addEventListener("change", (event) => onRowChange(state, event));
  rows.addEventListener("keydown", (event) => onRowKeyDown(state, event));
  rows.addEventListener("click", (event) => onRowClick(state, event));
  add.addEventListener("click", () => addColor(state));
  paste.addEventListener("paste", (event) => {
    const text = event.clipboardData?.getData("text");
    // No text handed over: the browser pastes into the field and change reads it.
    if (!text) return;
    event.preventDefault();
    applyText(state, text);
  });
  paste.addEventListener("keydown", (event) => {
    if (event.key === "Enter") takeText(state);
  });
  paste.addEventListener("change", () => takeText(state));

  container.classList.add("palette-editor");
  container.replaceChildren(rows, add, label, status);
  return state;
}

/**
 * Draws the editor into `container`, or brings it up to date when it already
 * holds one: call it again whenever the settings change. Builds with the
 * container's own document, so a parsed copy of the pane gets the same buttons.
 */
export function renderPaletteEditor(
  container: HTMLElement,
  options: PaletteEditorOptions,
): void {
  const state = states.get(container) ?? build(container, options);
  states.set(container, state);
  state.options = options;
  state.settings = options.settings;
  syncRows(state);
}

/**
 * Colors that arrive from outside the field (the Brand tab's Read from
 * selection), under the same rule as a paste. `dropped` counts the valid colors
 * the caller already left out; `none` is the sentence for an empty list.
 * Answers the sentence the status line now shows.
 */
export function applyBrandColors(
  container: HTMLElement,
  colors: readonly string[],
  dropped = 0,
  none = "No colors found.",
): string {
  const state = states.get(container);
  if (!state) return none;
  const list = colors.slice(0, MAX_BRAND_COLORS);
  const left = dropped + colors.length - list.length;
  return applyColors(state, list, left, none);
}

/** A sentence on the editor's status line, for pickers outside it that refuse in its voice. */
export function sayInEditor(container: HTMLElement, message: string): void {
  const state = states.get(container);
  if (state) say(state, message);
}
