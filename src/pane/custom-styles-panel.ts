// The Tools tab's Custom styles section: the name box and Save style, the list (apply, Update from the
// active cell, Delete behind a two-press confirm that counts the cells) and Starter styles from palette.
// DOM and wiring only (src/excel/custom-styles.ts writes the styles): one delegated click handler hands
// a row's press to the pane's guard, and the list is redrawn after every action and every palette save.

import { countText } from "../ui/count-text";
import {
  addStarterStyles,
  applyCustomStyle,
  countStyleWearers,
  deleteCustomStyle,
  listCustomStyles,
  saveCustomStyle,
  updateCustomStyle,
} from "../excel";
import { armConfirm, type ConfirmButton } from "../ui/confirm";
import { getElement } from "../ui/dom";
import { renderStyleRows, STYLE_ACTIONS } from "./custom-style-rows";
import { isExcelReady } from "./shared";

/** What a press on a row's button asks the pane to run: the action and the style's name. */
export type RunStyleAction = (action: string, name: string) => void;

const UNREADABLE = "Your styles could not be read from this workbook.";

// The busy latch of the pane (src/pane/shared.ts) disables every button, so
// the always-present Save button says whether an action is running.
const paneBusy = (): boolean =>
  getElement<HTMLButtonElement>("style-save").disabled;

// Two reads can be in flight (boot and a palette save): only the newest draws.
let latest = 0;

/** Reads the workbook's styles and draws them; a failed read says so and throws nothing. */
export async function refreshCustomStyles(): Promise<void> {
  // A page without the section (a test of another tab) has nothing to draw.
  if (
    !isExcelReady() ||
    document.getElementById("custom-style-list") === null
  ) {
    return;
  }
  latest += 1;
  const mine = latest;
  try {
    const result = await listCustomStyles();
    if (mine === latest) renderStyleRows(result, paneBusy());
  } catch {
    if (mine === latest) renderStyleRows({ styles: [], note: UNREADABLE });
  }
}

/**
 * Wires the name box (Enter presses Save style, so it runs through the same
 * guard) and the list's one click handler. Called once at boot; the static
 * buttons are wired by main.ts like every other [data-action] button.
 */
export function installCustomStylesPanel(run: RunStyleAction): void {
  const field = getElement<HTMLInputElement>("style-name");
  field.addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.isComposing) return;
    event.preventDefault();
    getElement<HTMLButtonElement>("style-save").click();
  });
  getElement("custom-style-list").addEventListener("click", (event) => {
    if (!(event.target instanceof Element)) return;
    const button = event.target.closest<HTMLButtonElement>(
      "button[data-action]",
    );
    const action = button?.dataset.action;
    const style = button?.dataset.style;
    if (action !== undefined && style !== undefined) run(action, style);
  });
}

// ---------------------------------------------------------------------------
// What each press runs (dispatch.ts calls these through the guard)
// ---------------------------------------------------------------------------

// The list is drawn again after the action, and also when it was refused: a
// style Excel's own gallery dropped is forgotten by the refusal and must leave
// the list with it.
async function thenRefresh(work: () => Promise<string>): Promise<string> {
  try {
    return await work();
  } finally {
    await refreshCustomStyles();
  }
}

/** Save style: the box's name, the active cell's look. The box clears once saved. */
export async function saveStyle(): Promise<string> {
  const field = getElement<HTMLInputElement>("style-name");
  const receipt = await thenRefresh(() => saveCustomStyle(field.value));
  field.value = "";
  return receipt;
}

export const applyStyle = (name: string): Promise<string> =>
  thenRefresh(() => applyCustomStyle(name));

export const updateStyle = (name: string): Promise<string> =>
  thenRefresh(() => updateCustomStyle(name));

export const addStarters = (): Promise<string> => thenRefresh(addStarterStyles);

// ---------------------------------------------------------------------------
// Delete: count, arm, then delete
// ---------------------------------------------------------------------------

// Keyed by the button, so a row drawn again starts unarmed. The count a press
// found rides on the button itself (data-cells) for the label to read.
const confirms = new WeakMap<HTMLButtonElement, ConfirmButton>();

function armedText(cells: number): string {
  if (cells === 0) return "Click again to delete";
  const what = cells === 1 ? "1 cell goes" : `${countText(cells)} cells go`;
  return `Click again: ${what} back to Normal`;
}

// No selector built from a style name: a name may hold any character.
function deleteButtonOf(name: string): HTMLButtonElement | null {
  const buttons = getElement(
    "custom-style-list",
  ).querySelectorAll<HTMLButtonElement>(
    `button[data-action="${STYLE_ACTIONS.remove}"]`,
  );
  return [...buttons].find((button) => button.dataset.style === name) ?? null;
}

function confirmOf(button: HTMLButtonElement): ConfirmButton {
  let held = confirms.get(button);
  if (!held) {
    held = armConfirm(button, () => undefined, {
      armedLabel: () => armedText(Number(button.dataset.cells ?? 0)),
    });
    confirms.set(button, held);
  }
  return held;
}

/**
 * Deleting sends every cell wearing the style back to Normal, so the first
 * press counts them and arms the button with that number; the second press, on
 * the same row within the confirm window, deletes.
 */
export async function deleteStyle(name: string): Promise<string> {
  const armed = deleteButtonOf(name);
  if (armed !== null && confirmOf(armed).isArmed()) {
    confirmOf(armed).disarm();
    return thenRefresh(() => deleteCustomStyle(name));
  }

  const cells = await countStyleWearers(name);
  // The list may have been drawn again while the cells were counted.
  const button = deleteButtonOf(name);
  if (button === null) return `${name} is no longer in the list.`;
  button.dataset.cells = String(cells);
  confirmOf(button).arm();
  return `${armedText(cells)}.`;
}
