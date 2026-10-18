// Takes the modeller to a tool Find a tool found and stops there: its tab, the
// control in view and in focus, a short mark on it. It never presses the control;
// a second Enter, now on the focused control, does that the normal way. DOM only.
// Invariant: going to a tool changes where you are, never the workbook.

export const FOUND_CLASS = "tool-found";
export const FOUND_MARK_MS = 2_000;

export interface GoDeps {
  root: Document;
  /** Shows the tab with this id, the pane's own tab strip doing it. */
  activate(tabId: string): void;
}

// The control marked now, if any: one at a time, and the mark goes the moment
// focus moves on or when its time is up.
let marked: { control: HTMLElement; timer: number } | null = null;

function clearMark(): void {
  if (marked === null) return;
  window.clearTimeout(marked.timer);
  marked.control.classList.remove(FOUND_CLASS);
  marked.control.removeEventListener("blur", clearMark);
  marked = null;
}

// The mark is the focus ring every control has, held for a moment: a click on a
// result gives focus by script, which draws no ring of its own.
function markFound(control: HTMLElement): void {
  clearMark();
  control.classList.add(FOUND_CLASS);
  control.addEventListener("blur", clearMark);
  marked = { control, timer: window.setTimeout(clearMark, FOUND_MARK_MS) };
}

// A template the Templates search is filtering out is hidden for that box's own
// query. Going there takes the query away, so the row is on the page to land on.
function showFiltered(control: HTMLElement, root: Document): void {
  if (control.closest("[data-filtered]") === null) return;
  const box = root.getElementById("template-search");
  if (!(box instanceof HTMLInputElement)) return;
  box.value = "";
  box.dispatchEvent(new Event("input"));
}

/**
 * Shows `tabId` (when the control sits on one), brings the control into view,
 * puts focus on it and marks it for a moment. Presses nothing.
 */
export function goToTool(
  control: HTMLButtonElement,
  tabId: string | null,
  deps: GoDeps,
): void {
  if (tabId !== null) deps.activate(tabId);
  showFiltered(control, deps.root);
  control.scrollIntoView?.({ block: "center" });
  control.focus({ preventScroll: true });
  markFound(control);
}
