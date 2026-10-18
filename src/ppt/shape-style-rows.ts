// The Shape styles list as DOM: a row per style (apply with a preview of its fill, outline
// and text, an Update icon, a Delete icon) or the line that says why there is none. No
// Office.js, so the help test and the UX gate draw it. Invariant: a row's buttons carry
// data-action and data-style, and a name only ever travels as text or an attribute value.

import type { ShapeStylePreview } from "../shape-styles";
import { getElement } from "../ui/dom";

export const SHAPE_STYLE_ACTIONS = {
  apply: "shape-style-apply",
  update: "shape-style-update",
  remove: "shape-style-delete",
} as const;

const EMPTY_NOTE = "No shape styles saved yet.";
// The preview is 22 pixels tall: an outline thicker than this would eat it.
const MIN_OUTLINE_PX = 1;
const MAX_OUTLINE_PX = 3;

/** The styles to draw, and why the list is not the whole story when it is not. */
export interface ShapeStyleList {
  styles: ShapeStylePreview[];
  note?: string;
}

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

// The style's own look, painted inline (the colors are the user's data): the fill
// behind "Aa" (drawn by the sheet), the outline as the sample's border and the text
// color on the letters. A part the style leaves alone or clears paints nothing; an
// outline that is not there takes the sample's hairline clear, so it never passes
// for one, and an outline of the theme's own color keeps the hairline.
function sampleOf(doc: Document, style: ShapeStylePreview): HTMLSpanElement {
  const sample = make(doc, "span", "cs-sample");
  sample.setAttribute("aria-hidden", "true");
  if (style.fill !== null) sample.style.backgroundColor = style.fill;
  if (style.textColor !== null) sample.style.color = style.textColor;
  sample.classList.toggle("is-bold", style.bold);
  const outline = style.outline;
  if (outline === null) {
    sample.style.borderColor = "transparent";
    return sample;
  }
  const px = Math.round(outline.weight);
  sample.style.borderWidth = `${String(Math.min(MAX_OUTLINE_PX, Math.max(MIN_OUTLINE_PX, px)))}px`;
  sample.style.borderStyle = outline.dashed ? "dashed" : "solid";
  if (outline.color !== null) sample.style.borderColor = outline.color;
  return sample;
}

function rowOf(
  doc: Document,
  style: ShapeStylePreview,
  busy: boolean,
): HTMLLIElement {
  const row = make(doc, "li", "cs-row");
  row.dataset.style = style.name;
  row.dataset.kind = style.kind;

  const apply = make(doc, "button", "cs-apply");
  apply.type = "button";
  apply.dataset.action = SHAPE_STYLE_ACTIONS.apply;
  apply.dataset.style = style.name;
  apply.title = `Put ${style.name} on the selected objects`;
  apply.append(sampleOf(doc, style), make(doc, "span", "cs-name", style.name));

  row.append(
    apply,
    iconButton(
      doc,
      "cs-update",
      "↻",
      SHAPE_STYLE_ACTIONS.update,
      style.name,
      "Update from the selected shape",
      `Update ${style.name} from the selected shape`,
    ),
    iconButton(
      doc,
      "cs-delete",
      "×",
      SHAPE_STYLE_ACTIONS.remove,
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
 * Draws one row per style into `list`, and into `note` the reason there are none
 * (the empty line, or why the list could not be read) or what was left out of the
 * list drawn. Builds with the list's own document, so a parsed copy of the pane
 * (the help coverage test) gets the buttons the live pane does.
 */
export function drawShapeStyleRows(
  list: HTMLElement,
  note: HTMLElement,
  result: ShapeStyleList,
  busy = false,
): void {
  const doc = list.ownerDocument;
  list.replaceChildren(
    ...result.styles.map((style) => rowOf(doc, style, busy)),
  );
  list.hidden = result.styles.length === 0;
  note.hidden = result.styles.length > 0 && result.note === undefined;
  note.textContent = result.note ?? EMPTY_NOTE;
}

/** The same into the pane's own list and empty line. */
export function renderShapeStyleRows(
  result: ShapeStyleList,
  busy = false,
): void {
  drawShapeStyleRows(
    getElement("shape-style-list"),
    getElement("shape-style-empty"),
    result,
    busy,
  );
}
