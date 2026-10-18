// What each press of the Shape styles section does: save the selected shape's look under a
// name, apply, update and delete a style, add the six starters, each against the user's list
// (shape-styles-store.ts) and the host (shape-style-io.ts). Each answers the receipt the
// pane toasts or throws one plain sentence. Invariant: a refusal writes nothing.

import type { KeyStore } from "../link/workspace";
import type { BrandSettings } from "../settings";
import {
  addShapeStyle,
  checkShapeStyleName,
  removeShapeStyle,
  replaceShapeLook,
  starterShapeStyles,
  withStarters,
  type ShapeStyle,
} from "../shape-styles";
import { startersDone, yoursNow } from "../starter-receipts";
import { captureShapeLook, paintShapeLook } from "./shape-style-io";
import { changeShapeStyles, readShapeStyles } from "./shape-styles-store";

export interface StyleFlowDeps {
  store: KeyStore;
  /** The palette the pane shows now: the starters are drawn from it. */
  palette(): BrandSettings;
  /** A line for the toast's details, beside the receipt a flow answers. */
  note(line: string): void;
}

const NOT_SOLID = "Shape styles currently support solid or no-fill objects.";
const SAVE_WORDS = {
  selectOne: "Select one object to save its style.",
  notSolid: NOT_SOLID,
};
const UPDATE_WORDS = {
  selectOne: "Select one object to update the style from.",
  notSolid: NOT_SOLID,
};
const APPLY_WORDS = { selectSome: "Select at least one object to style." };

const sameName = (first: string, second: string): boolean =>
  first.toLowerCase() === second.toLowerCase();
const gone = (name: string): string =>
  `The style ${name} is no longer in your list.`;
const objects = (count: number): string => (count === 1 ? "object" : "objects");

// What a list read said about itself, passed on once it was changed.
function say(deps: StyleFlowDeps, note: string | undefined): void {
  if (note !== undefined) deps.note(note);
}

/** Saves the selected shape's look as a style named `rawName`, last in the list. */
export async function saveShapeStyle(
  deps: StyleFlowDeps,
  rawName: string,
): Promise<string> {
  // The cheap refusals first: an empty or too long name never reaches PowerPoint.
  const typed = checkShapeStyleName(rawName, []);
  if (!typed.ok) throw new Error(typed.message);
  const look = await captureShapeLook(SAVE_WORDS);
  const { result, note } = await changeShapeStyles(deps.store, (styles) => {
    const named = checkShapeStyleName(
      typed.name,
      styles.map((style) => style.name),
    );
    if (!named.ok) throw new Error(named.message);
    const added = addShapeStyle(styles, {
      name: named.name,
      kind: "user",
      ...look,
    });
    if (!added.ok) throw new Error(added.message);
    return { styles: added.styles, result: named.name };
  });
  say(deps, note);
  return `Saved style ${result}`;
}

/** Puts the style on every selected shape. */
export async function applyShapeStyle(
  deps: StyleFlowDeps,
  name: string,
): Promise<string> {
  const read = await readShapeStyles(deps.store);
  if (!read.ok) throw new Error(read.message);
  const style = read.styles.find((held) => sameName(held.name, name));
  if (style === undefined) throw new Error(gone(name));
  const { painted, skipped } = await paintShapeLook(style, APPLY_WORDS);
  if (painted === 0) {
    throw new Error(
      `Style ${style.name} does not apply to the selected ${objects(skipped)}.`,
    );
  }
  const left = skipped > 0 ? `, ${String(skipped)} skipped` : "";
  return `Style ${style.name} on ${String(painted)} ${objects(painted)}${left}.`;
}

/**
 * Rewrites the style from the selected shape: it becomes what the shape looks like.
 * A starter updated this way becomes the user's own style (since 07.10.2026), so
 * the receipt says so and Starter styles from palette leaves it alone afterwards.
 */
export async function updateShapeStyle(
  deps: StyleFlowDeps,
  name: string,
): Promise<string> {
  const look = await captureShapeLook(UPDATE_WORDS);
  const { result, note } = await changeShapeStyles(deps.store, (styles) => {
    const held = styles.find((style) => sameName(style.name, name));
    const next = replaceShapeLook(styles, name, look);
    if (held === undefined || next === null) throw new Error(gone(name));
    return {
      styles: next,
      result: { name: held.name, wasStarter: held.kind === "starter" },
    };
  });
  say(deps, note);
  return result.wasStarter
    ? `Updated ${result.name}: it is yours now.`
    : `Updated style ${result.name}.`;
}

/** Takes the style out of the list. */
export async function deleteShapeStyle(
  deps: StyleFlowDeps,
  name: string,
): Promise<string> {
  const { result, note } = await changeShapeStyles(deps.store, (styles) => {
    const held = styles.find((style) => sameName(style.name, name));
    if (held === undefined) return { styles, result: null };
    return { styles: removeShapeStyle(styles, name), result: held.name };
  });
  say(deps, note);
  return result === null
    ? `The style ${name} was already gone.`
    : `Deleted style ${result}.`;
}

// ---------------------------------------------------------------------------
// Starter styles
// ---------------------------------------------------------------------------

// The shared wording (src/starter-receipts.ts) first, then what the cap left out.
function starterReceipt(
  added: number,
  restyled: number,
  leftAlone: string[],
  leftOut: string[],
): string {
  const alone = leftAlone.length === 0 ? "" : yoursNow(leftAlone);
  const out =
    leftOut.length === 0
      ? ""
      : ` Left out ${leftOut.join(", ")}: the list holds up to 30 styles.`;
  return startersDone(added, restyled) + alone + out;
}

/**
 * The six starters from the palette the pane shows now: the missing ones are
 * created, the ones pls,fix made are rewritten whole, so pressing it after a
 * palette change restyles them. A style the user holds under a starter's name,
 * saved by hand or a starter they updated, is left exactly as it is, no starter
 * is made under that name, and the receipt names it as theirs.
 */
export async function addStarterShapeStyles(
  deps: StyleFlowDeps,
): Promise<string> {
  const starters = starterShapeStyles(deps.palette());
  const { result, note } = await changeShapeStyles(deps.store, (styles) => {
    const merged = withStarters(styles, starters);
    return { styles: merged.styles, result: merged };
  });
  say(deps, note);
  return starterReceipt(
    result.added.length,
    result.restyled.length,
    result.leftAlone,
    result.leftOut,
  );
}

export type { ShapeStyle };
