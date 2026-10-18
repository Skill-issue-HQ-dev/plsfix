// The Custom styles list as DOM: one row per style (apply button with a preview, Update icon, Delete
// icon) or the line that says why there is none. No Office.js and no pane state, so the help test and
// the UX gate can draw it. Invariant: a row's buttons carry data-action and data-style, and a name
// only ever travels as text or as an attribute value.

import type { CustomStyleList } from "../excel";
import type { StylePreview } from "../custom-styles";
import { getElement } from "../ui/dom";

export const STYLE_ACTIONS = {
  apply: "style-apply",
  update: "style-update",
  remove: "style-delete",
} as const;
const EMPTY_NOTE = "No pls,fix styles in this workbook yet.";

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

function iconButton(
  doc: Document,
  className: string,
  glyph: string,
  action: string,
  name: string,
  label: string,
  title: string,
): HTMLButtonElement {
  const button = make(doc, "button", `icon-button ${className}`, glyph);
  button.type = "button";
  button.dataset.action = action;
  button.dataset.style = name;
  button.setAttribute("aria-label", label);
  button.title = title;
  return button;
}

// The style's own look, painted inline (the workbook's colors are data): the
// fill behind "Aa" (drawn by the sheet, so the button reads as the name alone),
// in the style's font color, bold when the style is. A style that leaves the
// fill or the font alone paints nothing there.
function sampleOf(doc: Document, style: StylePreview): HTMLSpanElement {
  const sample = make(doc, "span", "cs-sample");
  sample.setAttribute("aria-hidden", "true");
  if (style.fill !== null) sample.style.backgroundColor = style.fill;
  if (style.fontColor !== null) sample.style.color = style.fontColor;
  sample.classList.toggle("is-bold", style.bold);
  return sample;
}

function rowOf(
  doc: Document,
  style: StylePreview,
  busy: boolean,
): HTMLLIElement {
  const row = make(doc, "li", "cs-row");
  row.dataset.style = style.name;
  row.dataset.kind = style.kind;

  const apply = make(doc, "button", "cs-apply");
  apply.type = "button";
  apply.dataset.action = STYLE_ACTIONS.apply;
  apply.dataset.style = style.name;
  apply.title = `Put ${style.name} on the selected cells`;
  apply.append(sampleOf(doc, style), make(doc, "span", "cs-name", style.name));

  row.append(
    apply,
    iconButton(
      doc,
      "cs-update",
      "↻",
      STYLE_ACTIONS.update,
      style.name,
      "Update from the active cell",
      `Update ${style.name} from the active cell`,
    ),
    iconButton(
      doc,
      "cs-delete",
      "×",
      STYLE_ACTIONS.remove,
      style.name,
      "Delete",
      `Delete ${style.name}`,
    ),
  );
  // Rows born while an action runs follow the pane's busy latch, which only
  // reaches buttons that existed when it went on.
  for (const button of row.querySelectorAll("button")) button.disabled = busy;
  return row;
}

/**
 * Draws one row per style into `list`, or the reason there are none into `note`
 * (the empty line, or why the list could not be read). Builds with the list's
 * own document, so a parsed copy of the pane (the help coverage test) gets the
 * buttons the live pane does. The armed state of an old row goes with it.
 */
export function drawStyleRows(
  list: HTMLElement,
  note: HTMLElement,
  result: CustomStyleList,
  busy = false,
): void {
  const doc = list.ownerDocument;
  list.replaceChildren(
    ...result.styles.map((style) => rowOf(doc, style, busy)),
  );
  list.hidden = result.styles.length === 0;
  note.hidden = result.styles.length > 0;
  if (result.styles.length === 0) {
    note.textContent = result.note ?? EMPTY_NOTE;
  }
}

/** The same into the pane's own list and empty line. */
export function renderStyleRows(result: CustomStyleList, busy = false): void {
  drawStyleRows(
    getElement("custom-style-list"),
    getElement("custom-style-empty"),
    result,
    busy,
  );
}
