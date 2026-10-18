// The Templates picker: a search box over the registry and its templates as
// rows grouped by area, an Insert button on each carrying the entry's id as its
// data-action, so dispatch, help, tool search, click-through and sweep reach it
// like any button. DOM only: what a press does is the pane's `run`.

import {
  TEMPLATE_GROUPS,
  TEMPLATES,
  type TemplateEntry,
  type TemplateGroup,
} from "../templates/index";
import { filterTemplates } from "../templates/search";
import { getElement } from "../ui/dom";

/** What a press on an Insert button asks the pane to run: the entry's id. */
export type RunTemplate = (action: string) => void;

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

function rowNode(doc: Document, entry: TemplateEntry): HTMLLIElement {
  const row = make(doc, "li", "template-row");
  row.dataset.template = entry.id;
  const text = make(doc, "div", "template-text");
  text.append(
    make(doc, "strong", undefined, entry.name),
    make(doc, "span", undefined, entry.description),
  );
  const insert = make(doc, "button", "template-insert", "Insert");
  insert.type = "button";
  insert.dataset.action = entry.id;
  insert.setAttribute("aria-label", `Insert ${entry.name}`);
  // The tooltip carries what the row has no room for: the block's long
  // description and the size it writes.
  const { rows, cols, description } = entry.block;
  insert.title = `${description} Writes ${String(rows)} rows by ${String(cols)} columns at the active cell.`;
  row.append(text, insert);
  return row;
}

function groupNode(
  doc: Document,
  group: TemplateGroup,
  entries: readonly TemplateEntry[],
): HTMLElement {
  const node = make(doc, "div", "template-group");
  node.dataset.group = group;
  node.setAttribute("role", "group");
  const heading = make(doc, "h3", "template-group-name", group);
  heading.id = `template-group-${group.toLowerCase()}`;
  node.setAttribute("aria-labelledby", heading.id);
  const rows = make(doc, "ul", "template-rows");
  rows.append(...entries.map((entry) => rowNode(doc, entry)));
  node.append(heading, rows);
  return node;
}

/**
 * Fills `list` with one headed group per group that has a template, in the
 * picker's order, one row per entry. Builds with the list's own document, so a
 * parsed copy of the pane (the help coverage test) gets the same buttons the
 * live pane does.
 */
export function renderTemplateList(
  list: HTMLElement,
  entries: readonly TemplateEntry[],
): void {
  const doc = list.ownerDocument;
  list.replaceChildren(
    ...TEMPLATE_GROUPS.flatMap((group) => {
      const members = entries.filter((entry) => entry.group === group);
      return members.length === 0 ? [] : [groupNode(doc, group, members)];
    }),
  );
}

// A node the query takes out of view is `hidden` and carries `data-filtered`:
// the mark tells the tool search this hiding is only the picker's own filter,
// so a template this box is not showing is still one Find a tool may insert.
function setFiltered(node: HTMLElement, filtered: boolean): void {
  node.hidden = filtered;
  node.toggleAttribute("data-filtered", filtered);
}

// Rows outside the query are hidden, not removed: their buttons stay in the
// page for the tool search and the help card, and a group with no row left
// takes its heading with it.
function showMatches(
  list: HTMLElement,
  note: HTMLElement,
  entries: readonly TemplateEntry[],
  query: string,
): void {
  const keep = new Set(
    filterTemplates(entries, query).map((entry) => entry.id),
  );
  for (const row of list.querySelectorAll<HTMLElement>(".template-row")) {
    setFiltered(row, !keep.has(row.dataset.template ?? ""));
  }
  for (const group of list.querySelectorAll<HTMLElement>(".template-group")) {
    const empty = group.querySelector(".template-row:not([hidden])") === null;
    setFiltered(group, empty);
  }
  note.hidden = keep.size > 0;
}

/**
 * Builds the rows into `#template-list` and wires the search box, the Insert
 * buttons and the empty note. Called once, before the help copy reads the
 * section's buttons. A press, or Enter on the first match, goes through `run`.
 */
export function installTemplatePicker(
  run: RunTemplate,
  entries: readonly TemplateEntry[] = TEMPLATES,
): void {
  const search = getElement<HTMLInputElement>("template-search");
  const list = getElement<HTMLElement>("template-list");
  const note = getElement<HTMLElement>("template-empty");
  renderTemplateList(list, entries);
  const refresh = (): void => showMatches(list, note, entries, search.value);

  search.addEventListener("input", refresh);
  search.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      // An empty box has chosen nothing: Enter must not write the first row
      // of a list nobody narrowed.
      if (search.value.trim() === "") return;
      // A click, not a call: the same guard, the same busy latch, and a
      // disabled button (the pane is mid-action) answers nothing.
      const first = list.querySelector<HTMLButtonElement>(
        ".template-row:not([hidden]) button[data-action]",
      );
      if (first === null) return;
      event.preventDefault();
      first.click();
    } else if (event.key === "Escape" && search.value !== "") {
      event.preventDefault();
      search.value = "";
      refresh();
    }
  });
  list.addEventListener("click", (event) => {
    if (!(event.target instanceof Element)) return;
    const button = event.target.closest<HTMLButtonElement>(
      "button[data-action]",
    );
    if (button?.dataset.action !== undefined) run(button.dataset.action);
  });
}
