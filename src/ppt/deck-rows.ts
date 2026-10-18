// The Linked objects list as the deck sees it: the tags scan crossed with the
// deck's records, so a record no shape carries is a "Not on a slide" row. Pure.
// Invariant: only a settled scan may say a record has no shape - far slides
// read empty for a while after a deck opens (S1 spike, section 6).

import type { DeckRecord } from "../link/deck-record";
import type { LinkId, LinkKind } from "../link/model";
import { groupByProject } from "../link/project";
import type { LinkFilter } from "./actions";
import type { DrawFrom } from "./deck-drawable";

// What the row's status cell and the search box call a record with no shape.
export const OFF_SLIDE_WORD = "Not on a slide";

// A record with no shape, as the table draws it. `pushedAt` is the time of the
// push the record holds (the Sent column); `canPutBack` is false when the deck
// cannot draw the link again here (deck-drawable.ts), and then `deadWhy` is the
// row's sentence for it and `deadTag` the short word the row's second line says.
export interface DeckRowView {
  offSlide: true;
  key: string;
  id: LinkId;
  label: string;
  source: string;
  project?: string;
  kind: LinkKind;
  pushedAt: string;
  canPutBack: boolean;
  deadWhy?: string;
  deadTag?: string;
}

export function isDeckRow(row: object): row is DeckRowView {
  return "offSlide" in row && row.offSlide === true;
}

export function deckRowView(record: DeckRecord, draw: DrawFrom): DeckRowView {
  const drawing = draw(record);
  const dead = drawing.ok
    ? {}
    : {
        deadWhy: `${upFirst(drawing.issue)}. ${drawing.way}`,
        deadTag: drawing.tag,
      };
  return {
    offSlide: true,
    key: `deck/${record.id}`,
    id: record.id,
    label: record.label,
    source: record.src.workbook,
    ...(record.project === undefined ? {} : { project: record.project }),
    kind: record.kind,
    pushedAt: record.pushedAt,
    canPutBack: drawing.ok,
    ...dead,
  };
}

function upFirst(text: string): string {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

// The records of the deck that no shape carries, as rows, label order. An
// unsettled scan has no say: it answers nothing, however empty it looks. A
// link on many slides is carried once; the shape rows are the shapes' own, so
// a record with a shape adds nothing here.
export function offSlideRows(
  found: readonly { tag: { id: LinkId } }[],
  records: ReadonlyMap<LinkId, DeckRecord>,
  settled: boolean,
  draw: DrawFrom,
): DeckRowView[] {
  if (!settled) return [];
  const carried = new Set(found.map((one) => one.tag.id));
  return [...records.values()]
    .filter((record) => !carried.has(record.id))
    .map((record) => deckRowView(record, draw))
    .sort((a, b) => a.label.localeCompare(b.label) || a.id.localeCompare(b.id));
}

// The list's own filters, applied to a row that has no slide and no relay
// status: a slide or a status filter shows none of them, the rest follow the
// record. The search reaches the words "Not on a slide" too.
export function filterDeckRows(
  rows: readonly DeckRowView[],
  filter: LinkFilter,
): DeckRowView[] {
  if (filter.status !== "all") return [];
  if (filter.slide !== undefined && filter.slide !== "all") return [];
  const query = filter.query.trim().toLocaleLowerCase();
  return rows.filter((row) => {
    if (
      filter.source !== undefined &&
      filter.source !== "all" &&
      row.source !== filter.source
    ) {
      return false;
    }
    if (
      filter.project !== undefined &&
      filter.project !== "all" &&
      (row.project ?? "") !== filter.project
    ) {
      return false;
    }
    if (query === "") return true;
    const values = [
      row.label,
      row.source,
      row.project ?? "",
      row.kind,
      OFF_SLIDE_WORD,
    ];
    return values.some((value) => value.toLocaleLowerCase().includes(query));
  });
}

// The one grouping rule of the deck list - groupByProject - over the shapes'
// rows and then the records': inside a folder the shape rows keep their order
// and the rows with no shape follow them.
export function unionByProject<T extends { project?: string }>(
  shapes: readonly T[],
  deck: readonly DeckRowView[],
): [string, (T | DeckRowView)[]][] {
  const rows: (T | DeckRowView)[] = [...shapes, ...deck];
  return groupByProject(rows, (row) => row.project);
}
