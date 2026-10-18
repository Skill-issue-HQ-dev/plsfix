// The user's list of shape styles in the Office key store, one key, so it is there in every
// deck on this computer. Owns reading it (empty, readable, corrupt, from a newer pls,fix,
// refused by the host) and the one step a change takes: read, change, write. Invariant: a
// list that could not be read is unknown, never empty: nothing is built or written on it.

import type { KeyStore } from "../link/workspace";
import type { ShapeStyle } from "../shape-styles";
import {
  decodeShapeStyles,
  encodeShapeStyles,
  SHAPE_STYLES_KEY,
} from "../shape-styles-codec";

export type StoredStyles =
  /** `note` says why the list is shorter than the store holds. */
  | { ok: true; styles: ShapeStyle[]; note?: string }
  /** A newer pls,fix wrote it, or the store would not answer: no list at all. */
  | { ok: false; message: string; why: "newer" | "unreadable" };

const UNREAD =
  "Your saved shape styles could not be read, so nothing was changed.";

export async function readShapeStyles(store: KeyStore): Promise<StoredStyles> {
  let raw: string | null;
  try {
    raw = await store.get(SHAPE_STYLES_KEY);
  } catch {
    return { ok: false, message: UNREAD, why: "unreadable" };
  }
  const read = decodeShapeStyles(raw);
  if (!read.ok) return { ...read, why: "newer" };
  return read.reason === undefined
    ? { ok: true, styles: read.styles }
    : { ok: true, styles: read.styles, note: read.reason };
}

/** Stores the list; says in its own words that it was not saved when the store refuses. */
export async function writeShapeStyles(
  store: KeyStore,
  styles: readonly ShapeStyle[],
): Promise<void> {
  try {
    await store.set(SHAPE_STYLES_KEY, encodeShapeStyles(styles));
  } catch (error) {
    const why = error instanceof Error ? error.message : String(error);
    throw new Error(`Your shape styles were not saved: ${why}`);
  }
}

/**
 * One change as one step: the list is read fresh, `change` makes the new one (it
 * throws a sentence to refuse, and then nothing is written; handing back the very
 * list it was given writes nothing either), and the new list is stored. Answers
 * what the change answered and the note of a list that was read with entries left
 * out. A list that cannot be read, or that a newer pls,fix wrote, is refused.
 */
export async function changeShapeStyles<T>(
  store: KeyStore,
  change: (styles: ShapeStyle[]) => { styles: ShapeStyle[]; result: T },
): Promise<{ result: T; note?: string }> {
  const read = await readShapeStyles(store);
  if (!read.ok) throw new Error(read.message);
  const changed = change(read.styles);
  if (changed.styles !== read.styles) {
    await writeShapeStyles(store, changed.styles);
  }
  return { result: changed.result, note: read.note };
}
