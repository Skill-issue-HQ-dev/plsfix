// The "Not on a slide" row of the Linked objects table: a record the deck holds
// for an object no slide carries, with the button that brings it back and the
// one that lets the record go (two presses).
// No tick (the toolbar acts on shapes) and no slide number; Put back is dead,
// with the reason, when the deck cannot draw the link again here.

import type { LinkId } from "../link/model";
import { armConfirm } from "../ui/confirm";
import { relativeStamp } from "../ui/time";
import { OFF_SLIDE_WORD, type DeckRowView } from "./deck-rows";
import { emptyCell, labelCell, sourceMeta, textCell } from "./view-cells";

// The records the deck holds for objects no slide carries, and what the pane
// does when one of their buttons is pressed. Passed beside the shapes' rows:
// they fall into the same project folders (deck-rows.ts unionByProject).
export interface DeckRowsPanel {
  rows: DeckRowView[];
  onPutBack: (id: LinkId) => void;
  onRemove: (id: LinkId) => void;
}

const HOLDS_BADGE = `${OFF_SLIDE_WORD}: deleted from the slides, but this deck still holds its link and data.`;

export function offSlideRow(
  row: DeckRowView,
  deck: DeckRowsPanel | undefined,
): HTMLTableRowElement {
  const line = document.createElement("tr");
  line.className = "link-offslide";
  line.dataset.deckId = row.id;
  const label = labelCell(
    row.label,
    sourceMeta(row.source, row.kind),
    deckActions(row, deck),
  );
  // Why Put back is dead is said on the row, not only in a tooltip a disabled
  // button may never show: a line of its own under the source.
  if (!row.canPutBack) label.querySelector(".link-meta")!.after(deadLine(row));
  line.append(
    emptyCell("link-select"),
    textCell("link-slide", "-"),
    label,
    offSlideStatus(row),
    textCell("link-updated", relativeStamp(row.pushedAt)),
  );
  return line;
}

function deadLine(row: DeckRowView): HTMLElement {
  const line = document.createElement("span");
  line.className = "link-meta link-dead";
  line.textContent = row.deadTag ?? "";
  line.title = row.deadWhy ?? "";
  return line;
}

function offSlideStatus(row: DeckRowView): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-status";
  const badge = document.createElement("span");
  badge.className = "badge offslide";
  badge.textContent = OFF_SLIDE_WORD;
  const detail = row.canPutBack
    ? HOLDS_BADGE
    : `${OFF_SLIDE_WORD}: deleted from the slides. ${row.deadWhy ?? ""}`;
  badge.title = detail;
  badge.setAttribute("aria-label", detail);
  cell.append(badge);
  return cell;
}

function deckActions(
  row: DeckRowView,
  deck: DeckRowsPanel | undefined,
): HTMLElement {
  const box = document.createElement("span");
  box.className = "deck-row-actions";
  const put = deckButton("deck-put-back", "Put back", `Put back ${row.label}`);
  put.title = row.canPutBack
    ? "Put it back at the slide and spot picked on the Inbox tab"
    : (row.deadWhy ?? "");
  if (!row.canPutBack) {
    put.disabled = true;
    // The pane's busy latch hands every button back enabled when an action
    // ends; this mark is what keeps a dead one dead (main.ts setBusy).
    put.dataset.dead = "true";
  }
  put.addEventListener("click", () => deck?.onPutBack(row.id));
  // Destructive and one-way, so the first press arms and the second removes.
  // Alive on a record that kept the identity only too: its key is what goes.
  const remove = deckButton(
    "deck-remove",
    "Remove from this deck",
    `Remove ${row.label} from this deck`,
  );
  remove.title = "Delete this deck's copy of the link and its key. Press twice";
  const confirm = armConfirm(remove, () => deck?.onRemove(row.id));
  remove.addEventListener("click", confirm.handleClick);
  box.append(put, remove);
  return box;
}

function deckButton(
  action: string,
  text: string,
  name: string,
): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "deck-row-action";
  button.dataset.action = action;
  button.textContent = text;
  button.setAttribute("aria-label", name);
  return button;
}
