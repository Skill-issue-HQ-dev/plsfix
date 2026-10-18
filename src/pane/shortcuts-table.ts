// Whether Office loaded pls,fix's key table: an add-in uploaded to Excel for the web
// never gets it (getShortcuts answers {}, replaceShortcuts rejects: no key works).
// Owns that verdict, its two sentences and Apply and Reset all's on/off state; the
// section's own flow is shortcuts-panel.ts.

/** What the section's note says instead of its hint, while the table is not loaded. */
export const NOT_LOADED =
  "Office has not loaded pls,fix's keyboard shortcuts here, so the keys do nothing in this window. The buttons and the ribbon work as usual.";

/** The one sentence for a replaceShortcuts that Office refused as a GeneralException. */
export const REFUSED = "Office refused the new keys, so nothing was changed.";

const BUTTONS = ["shortcuts-apply", "shortcuts-reset"];

let notLoaded = false;
let section: ParentNode | null = null;

/** A new section: nothing is known about the table yet. */
export function startTable(root: ParentNode): void {
  section = root;
  notLoaded = false;
}

/**
 * A getShortcuts answer that landed: the table is not loaded when the answer
 * holds none of the actions the panel lists. An answer that holds some is a
 * loaded table, a null among them (a conflict the user resolved answers null
 * for its action); a read that failed says nothing and never reaches here.
 */
export function readTable(
  answer: Record<string, string | null>,
  ids: string[],
): void {
  notLoaded = ids.length > 0 && !ids.some((id) => id in answer);
}

export function tableNotLoaded(): boolean {
  return notLoaded;
}

/**
 * Apply and Reset all are off while the table is not loaded. The busy latch
 * hands every button back when an action ends, so main.ts's rule-button sync
 * calls this to put these two back where the verdict left them.
 */
export function syncShortcutButtons(): void {
  for (const id of BUTTONS) {
    const button = section?.querySelector<HTMLButtonElement>(`#${id}`);
    if (button) button.disabled = notLoaded;
  }
}
