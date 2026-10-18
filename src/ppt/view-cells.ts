// The table cells the PowerPoint pane's Linked objects rows are built from,
// shared by a shape's row (views.ts) and a Not on a slide row (views-deck-row.ts).
// Pure DOM builders; the workbook-and-kind line is the same one the Inbox uses.

import { kindWord } from "../link/kind-words";
import type { LinkKind, LinkTag } from "../link/model";

// The workbook a link came from and what it is: the same cells can be exported
// twice, once as a picture and once as a table, and the kind is what tells the
// two apart in a list. `choice`: the shape's tag says the user kept a picture.
export function sourceMeta(
  workbook: string,
  kind: LinkKind,
  choice?: LinkTag["choice"],
): string {
  return `${workbook} · ${kindWord(kind, choice)}`;
}

// The title carries the full text: these columns are narrow enough in a task
// pane that a workbook name is usually ellipsised.
export function textCell(
  className: string,
  text: string,
): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = className;
  cell.textContent = text;
  cell.title = text;
  return cell;
}

export function emptyCell(className: string): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = className;
  return cell;
}

// The object name plus, on its own line, the workbook it comes from and its
// kind (sourceMeta): the deck list's only place for the source now that the
// Source column is gone. A row with buttons of its own carries them under the
// meta line.
export function labelCell(
  label: string,
  meta: string,
  actions?: HTMLElement,
): HTMLTableCellElement {
  const cell = document.createElement("td");
  cell.className = "link-label";
  const metaLine = document.createElement("span");
  metaLine.className = "link-meta";
  metaLine.textContent = meta;
  cell.append(document.createTextNode(label), metaLine);
  if (actions) cell.append(actions);
  cell.title = `${label} · ${meta}`;
  return cell;
}
