// The PowerPoint pane's DOM renderers: the link table body, the inbox list,
// the "Change source" candidate list and the labels they read. Pure - each
// function is handed its container, what to draw and the callback to call,
// touches nothing else, and clears the container first so a re-render can
// never leave a stale listener.

import type { InboxItem, LinkKind, LinkTag } from "../link/model";
import { groupByProject } from "../link/project";
import type { LinkStatus } from "../link/status";
import type { LinkTransport } from "../link/transport-setting";
import { clockStamp, NEVER, relativeStamp, relativeTime } from "../ui/time";
import { isDeckRow, unionByProject } from "./deck-rows";
import { emptyInboxLine } from "./mode-text";
import { labelCell, sourceMeta, textCell } from "./view-cells";
import { offSlideRow, type DeckRowsPanel } from "./views-deck-row";

export type { DeckRowsPanel } from "./views-deck-row";

export interface LinkRowView {
  key: string;
  slide: number;
  label: string;
  source: string;
  project?: string;
  kind: LinkKind;
  // The shape's tag says the user kept this chart as a picture: the row reads
  // "Picture (your choice)" instead of the kind's own word.
  choice?: LinkTag["choice"];
  status: LinkStatus;
  // null only for a link the relay has never held a push for; 0 is a real
  // (if ancient) time.
  pushedAt: number | null;
  selected: boolean;
}

const STATUS_LABELS: Record<LinkStatus, string> = {
  current: "Up to date",
  updateAvailable: "Update available",
  missing: "Not on the relay",
  wrongKey: "Wrong link key",
  notPasted: "Not pasted yet",
  inDeck: "In this deck",
};

// The whole sentence behind a badge that is too short to hold it: the tooltip
// and the accessible name. The relay answers 404 for a link it forgot after 30
// days without a push and for one the modeller removed in Excel, so a link it
// no longer holds is not a source that went missing.
const STATUS_DETAILS: Partial<Record<LinkStatus, string>> = {
  missing:
    "Not on the relay: expired after 30 days without a push, or removed in Excel. Push it again, or export it again if it was removed.",
};

// One badge class per status so the colour is CSS's business, not the renderer's.
const STATUS_CLASSES: Record<LinkStatus, string> = {
  current: "badge current",
  updateAvailable: "badge update",
  missing: "badge missing",
  wrongKey: "badge wrong",
  notPasted: "badge notpasted",
  inDeck: "badge indeck",
};

export function statusLabel(status: LinkStatus): string {
  return STATUS_LABELS[status];
}

// What a row the deck answered says about its data: when it was sent, as the
// house stamp (src/ui/time.ts). A deck whose record carries no usable time
// leaves the clause out rather than guess.
function dataFrom(dataAt: number | null | undefined): string | null {
  return dataAt === null || dataAt === undefined
    ? null
    : `data from ${clockStamp(dataAt)}`;
}

const IN_DECK_DETAIL =
  "No copy of this link was found on this computer or the relay, so the deck's own copy is shown. Copy or push it again from Excel to update.";

// `dataAt` (unix seconds) is the time the row's data was sent; only the
// In this deck sentence reads it.
export function statusDetail(
  status: LinkStatus,
  dataAt?: number | null,
): string {
  if (status === "inDeck") {
    const when = dataFrom(dataAt);
    return `${STATUS_LABELS.inDeck}${when === null ? "" : `, ${when}`}. ${IN_DECK_DETAIL}`;
  }
  return STATUS_DETAILS[status] ?? STATUS_LABELS[status];
}

// The table's column count: select, slide, object, status, updated. A folder
// header spans all five.
const LINK_TABLE_COLUMNS = 5;

export function renderLinkRows(
  table: HTMLTableSectionElement,
  rows: LinkRowView[],
  onToggle: (key: string, selected: boolean) => void,
  deck?: DeckRowsPanel,
): void {
  const lines: HTMLTableRowElement[] = [];
  for (const [name, group] of unionByProject(rows, deck?.rows ?? [])) {
    lines.push(folderHeader(name, group.length));
    for (const row of group) {
      lines.push(
        isDeckRow(row) ? offSlideRow(row, deck) : linkRow(row, onToggle),
      );
    }
  }
  table.replaceChildren(...lines);
}

// No checkbox, no data-key: a folder header is not a link row, just the
// project name and how many of the rows shown belong to it.
function folderHeader(name: string, count: number): HTMLTableRowElement {
  const line = document.createElement("tr");
  line.className = "link-folder";
  const cell = document.createElement("td");
  cell.colSpan = LINK_TABLE_COLUMNS;
  cell.textContent = `${name} · ${count}`;
  line.append(cell);
  return line;
}

function linkRow(
  row: LinkRowView,
  onToggle: (key: string, selected: boolean) => void,
): HTMLTableRowElement {
  const line = document.createElement("tr");
  line.dataset.key = row.key;
  line.append(
    selectCell(row, onToggle),
    textCell("link-slide", String(row.slide)),
    labelCell(row.label, sourceMeta(row.source, row.kind, row.choice)),
    statusCell(row),
    textCell("link-updated", relativeTime(row.pushedAt)),
  );
  return line;
}

function selectCell(
  row: LinkRowView,
  onToggle: (key: string, selected: boolean) => void,
): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-select";
  const box = document.createElement("input");
  box.type = "checkbox";
  box.checked = row.selected;
  box.setAttribute("aria-label", `Select ${row.label}`);
  box.addEventListener("change", () => onToggle(row.key, box.checked));
  cell.append(box);
  return cell;
}

function statusCell(row: LinkRowView): HTMLTableCellElement {
  const { status } = row;
  const cell = document.createElement("td");
  cell.className = "link-status";
  const badge = document.createElement("span");
  badge.className = STATUS_CLASSES[status];
  badge.textContent = statusLabel(status);
  // Narrow panes truncate the badge text; the title tooltip (and an
  // aria-label, since the badge has no other accessible name) keep the
  // full status readable/announced.
  const detail = statusDetail(status, row.pushedAt);
  badge.title = detail;
  badge.setAttribute("aria-label", detail);
  cell.append(badge);
  // A row the deck answered says which data it holds, under the word.
  const when = status === "inDeck" ? dataFrom(row.pushedAt) : null;
  if (when !== null) {
    const line = document.createElement("span");
    line.className = "link-meta";
    line.textContent = when;
    cell.append(line);
  }
  return cell;
}

// `mode` words the empty line (relay's when the caller names none).
export function renderInbox(
  list: HTMLElement,
  items: InboxItem[],
  onInsert: (item: InboxItem) => void,
  mode: LinkTransport = "relay",
): void {
  if (items.length === 0) {
    list.replaceChildren(hint(emptyInboxLine(mode)));
    return;
  }
  list.replaceChildren(...groupedInbox(items, onInsert));
}

function groupedInbox(
  items: InboxItem[],
  onInsert: (item: InboxItem) => void,
): HTMLElement[] {
  const nodes: HTMLElement[] = [];
  for (const [name, group] of groupByProject(items, (item) => item.project)) {
    nodes.push(inboxHeader(name));
    for (const item of group) nodes.push(inboxRow(item, onInsert));
  }
  return nodes;
}

function inboxHeader(name: string): HTMLParagraphElement {
  const line = document.createElement("p");
  line.className = "inbox-project";
  line.textContent = name;
  return line;
}

function hint(text: string): HTMLParagraphElement {
  const line = document.createElement("p");
  line.className = "hint";
  line.textContent = text;
  return line;
}

function inboxRow(
  item: InboxItem,
  onInsert: (item: InboxItem) => void,
): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "inbox-row";

  const label = document.createElement("strong");
  label.textContent = item.label;
  const meta = document.createElement("small");
  meta.textContent = inboxMeta(item);
  const text = document.createElement("div");
  text.className = "inbox-text";
  text.append(label, meta);

  const insert = document.createElement("button");
  insert.type = "button";
  insert.className = "inbox-insert";
  insert.textContent = "Insert";
  insert.setAttribute("aria-label", `Insert ${item.label}`);
  insert.addEventListener("click", () => onInsert(item));

  row.append(text, insert);
  return row;
}

// That, plus how long the export has waited: the label above already names the
// sheet and range.
function inboxMeta(item: InboxItem): string {
  const age = relativeStamp(item.createdAt);
  const text = sourceMeta(item.src.workbook, item.kind);
  return age === NEVER ? text : `${text} · ${age}`;
}

// The inbox's Slide picker: "This slide" (value "", the default meaning "the
// active one") plus one option per slide in the deck, by position - the
// picker only ever learns a count, never the slides' own ids (target.ts
// resolves the chosen position back to one when an insert actually runs).
// The current pick survives a re-render when it is still in range, and falls
// back to "This slide" the moment the deck holds fewer slides than that.
export function renderSlideOptions(
  select: HTMLSelectElement,
  count: number,
): void {
  const picked = select.value;
  const options = [new Option("This slide", "")];
  for (let index = 1; index <= count; index += 1) {
    options.push(new Option(`Slide ${String(index)}`, String(index)));
  }
  select.replaceChildren(...options);
  select.value = options.some((option) => option.value === picked)
    ? picked
    : "";
}

// The "Change source" chooser: every waiting export, best match first (the
// order is candidatesFor's, not this renderer's). The workbook and the age are
// what tell two exports of the same table apart, so both are on the line.
export function renderCandidates(
  select: HTMLSelectElement,
  items: InboxItem[],
): void {
  select.replaceChildren(...items.map(candidateOption));
}

function candidateOption(item: InboxItem): HTMLOptionElement {
  const option = document.createElement("option");
  option.value = item.id;
  const text = `${item.label} · ${inboxMeta(item)}`;
  option.textContent = text;
  option.title = text;
  return option;
}
