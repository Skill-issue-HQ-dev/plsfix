// The Tools tab's Shape styles section: the name box and Save style, the list (apply, Update,
// Delete behind a two-press confirm) and Starter styles from palette. DOM and wiring only: the
// presses are shape-styles-flows.ts, the rows shape-style-rows.ts, the list shape-styles-store.ts.
// Invariant: every press runs through the pane's guard, and the list is drawn again after it.

import type { KeyStore } from "../link/workspace";
import type { BrandSettings } from "../settings";
import { previewOf } from "../shape-styles";
import { armConfirm, type ConfirmButton } from "../ui/confirm";
import { getElement } from "../ui/dom";
import {
  renderShapeStyleRows,
  SHAPE_STYLE_ACTIONS,
  type ShapeStyleList,
} from "./shape-style-rows";
import {
  addStarterShapeStyles,
  applyShapeStyle,
  deleteShapeStyle,
  saveShapeStyle,
  updateShapeStyle,
  type StyleFlowDeps,
} from "./shape-styles-flows";
import { readShapeStyles } from "./shape-styles-store";

export interface ShapeStylesDeps {
  /** The pane's guarded runner: busy state, toast, error reporting. */
  act(run: () => Promise<string>, action: string): void;
  /** True while a guarded action runs: what the rows it draws are born as. */
  busy(): boolean;
  /** A line for the toast's details, beside the receipt a flow answers. */
  note(line: string): void;
  keyStore: KeyStore;
  /** The palette the pane shows: the starters are drawn from it. */
  palette(): BrandSettings;
}

export interface ShapeStyles {
  /** Reads the list and draws it; a failed read says so and throws nothing. */
  refresh(): Promise<void>;
  /** Save style: the box's name, the selected shape's look. The box clears once saved. */
  save(): Promise<string>;
  /** Starter styles from palette. */
  starters(): Promise<string>;
}

const UNREADABLE = "Your saved shape styles could not be read.";

// What the list shows for what the store said: the styles, or the one sentence of
// why there are none.
async function listOf(store: KeyStore): Promise<ShapeStyleList> {
  const read = await readShapeStyles(store);
  if (read.ok) return { styles: read.styles.map(previewOf), note: read.note };
  return {
    styles: [],
    note: read.why === "unreadable" ? UNREADABLE : read.message,
  };
}

// What the pieces below share: the deps, what the flows are given, the counter that
// lets only the newest read draw, and the armed Delete buttons.
interface Section {
  deps: ShapeStylesDeps;
  flow: StyleFlowDeps;
  /** Two reads can be in flight (the boot and an action): only the newest draws. */
  latest: number;
  /** Keyed by the button, so a row drawn again starts unarmed. */
  confirms: WeakMap<HTMLButtonElement, ConfirmButton>;
}

// What a row's press runs, by the row button's action.
const ROW_FLOWS: Record<
  string,
  (flow: StyleFlowDeps, name: string) => Promise<string>
> = {
  [SHAPE_STYLE_ACTIONS.apply]: applyShapeStyle,
  [SHAPE_STYLE_ACTIONS.update]: updateShapeStyle,
  [SHAPE_STYLE_ACTIONS.remove]: deleteShapeStyle,
};

async function refreshOf(section: Section): Promise<void> {
  section.latest += 1;
  const mine = section.latest;
  let drawn: ShapeStyleList;
  try {
    drawn = await listOf(section.deps.keyStore);
  } catch {
    drawn = { styles: [], note: UNREADABLE };
  }
  if (mine === section.latest) {
    renderShapeStyleRows(drawn, section.deps.busy());
  }
}

// The list is drawn again after the action, and also when it was refused: a style
// another deck's pane deleted is forgotten by the refusal.
async function thenRefresh(
  section: Section,
  work: () => Promise<string>,
): Promise<string> {
  try {
    return await work();
  } finally {
    await refreshOf(section);
  }
}

function runRow(section: Section, action: string, name: string): void {
  const flowOf = ROW_FLOWS[action];
  if (flowOf === undefined) return;
  section.deps.act(
    () => thenRefresh(section, () => flowOf(section.flow, name)),
    action,
  );
}

function confirmOf(
  section: Section,
  button: HTMLButtonElement,
  name: string,
): ConfirmButton {
  let held = section.confirms.get(button);
  if (!held) {
    held = armConfirm(button, () => {
      runRow(section, SHAPE_STYLE_ACTIONS.remove, name);
    });
    section.confirms.set(button, held);
  }
  return held;
}

function onListClick(section: Section, event: Event): void {
  if (!(event.target instanceof Element)) return;
  const button = event.target.closest<HTMLButtonElement>("button[data-action]");
  const action = button?.dataset.action;
  const name = button?.dataset.style;
  if (button == null || action === undefined || name === undefined) return;
  if (action === SHAPE_STYLE_ACTIONS.remove) {
    confirmOf(section, button, name).handleClick();
  } else runRow(section, action, name);
}

async function saveOf(section: Section): Promise<string> {
  const field = getElement<HTMLInputElement>("shape-style-name");
  const receipt = await thenRefresh(section, () =>
    saveShapeStyle(section.flow, field.value),
  );
  field.value = "";
  return receipt;
}

// Enter in the name box presses Save style, so it runs through the same guard.
function wire(section: Section): void {
  getElement("shape-style-name").addEventListener("keydown", (event) => {
    if (event.key !== "Enter" || event.isComposing) return;
    event.preventDefault();
    getElement<HTMLButtonElement>("save-shape-style").click();
  });
  getElement("shape-style-list").addEventListener("click", (event) => {
    onListClick(section, event);
  });
  // Another deck's pane may have saved or deleted a style since this one drew.
  getElement("tab-tools").addEventListener("click", () => {
    if (!section.deps.busy()) void refreshOf(section);
  });
}

/**
 * Wires the name box, the list's one click handler and the Tools tab's refresh. The
 * two static buttons are wired by main.ts like every other button, to `save` and
 * `starters`.
 */
export function installShapeStyles(deps: ShapeStylesDeps): ShapeStyles {
  const section: Section = {
    deps,
    flow: { store: deps.keyStore, palette: deps.palette, note: deps.note },
    latest: 0,
    confirms: new WeakMap(),
  };
  wire(section);
  return {
    refresh: () => refreshOf(section),
    save: () => saveOf(section),
    starters: () =>
      thenRefresh(section, () => addStarterShapeStyles(section.flow)),
  };
}
