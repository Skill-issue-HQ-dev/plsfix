// The Inbox as this pane sees it: the relay's rows, last posted first, minus
// what the deck already has, and which of them "Paste latest linked" means.
// Pure - no relay, no Office.js; links.ts does the fetching and the inserting
// and hands the deck's own tags in through a scan function.
// Invariants: the deck decides. A row for a link whose object sits on a slide is
// hidden (Excel's Push announces every link it pushes, so such rows pile up),
// a row for an object since deleted is offered again unless the deck's record
// of it can be drawn again on this host - the Linked objects list shows that
// one as Not on a slide and puts it back from there; a record it cannot draw
// (no data, or a chart this host draws as a picture only) leaves the row as
// the way back - and only a deck that could not be read falls back to this
// pane's memory of what it pasted. The list order is the
// relay's row time, which every Push renews; the one-click paste goes by each
// item's own export time instead.

import type { InboxItem } from "../link/model";

const pasted = new Set<string>();
const NO_IDS: ReadonlySet<string> = new Set();

export function rememberPasted(id: string): void {
  pasted.add(id);
}

// The relay orders rows by when each was last posted, newest first, ties broken
// by arrival, and a self-hosted one is free to answer in storage order, so the
// sort is repeated here; it is stable, so equal row times keep the server's
// order. A re-announce renews that time: this is "pushed last", not "exported
// last", which is why latestInboxItem does not simply take the first row.
// `deckLinks` is every link id the deck's tags name; undefined means the deck
// could not be read. With it, an export this pane pasted whose object has since
// left every slide is forgotten (its row may be Excel re-announcing it), so the
// two sets never disagree; without it, the pasted set alone keeps a delete the
// relay refused from turning the next one-click paste into a duplicate.
// `recorded` is every link id whose record in the deck can be drawn again here
// (deck-drawable.ts, read apart from the tags): hidden whatever the tags say.
export function queueOrder(
  rows: { item: InboxItem; createdAt: number }[],
  deckLinks?: ReadonlySet<string>,
  recorded: ReadonlySet<string> = NO_IDS,
): InboxItem[] {
  if (deckLinks !== undefined) {
    for (const id of pasted) if (!deckLinks.has(id)) pasted.delete(id);
  }
  const hidden = deckLinks ?? pasted;
  return rows
    .sort((left, right) => right.createdAt - left.createdAt)
    .map(({ item }) => item)
    .filter((item) => !hidden.has(item.id) && !recorded.has(item.id));
}

// What links.ts lists: the rows minus what the deck holds, read through `scan`
// (the host's tags scan). Nothing waiting costs no scan; a scan the host
// refuses leaves the deck unread rather than the Inbox empty.
export async function shownInbox(
  rows: { item: InboxItem; createdAt: number }[],
  scan: () => Promise<{ tag: { id: string } }[]>,
  recorded?: ReadonlySet<string>,
): Promise<InboxItem[]> {
  if (rows.length === 0) return [];
  let deckLinks: Set<string> | undefined;
  try {
    deckLinks = new Set((await scan()).map((found) => found.tag.id));
  } catch {
    deckLinks = undefined;
  }
  return queueOrder(rows, deckLinks, recorded);
}

// What "Paste latest linked" inserts: the newest EXPORT among the rows shown. An
// item's createdAt is the link's own creation time, stamped once in Excel and
// carried unchanged by every re-announce, while the relay's row time (the list
// order) moves on every Push; following the list would hand a week-old link
// pushed a minute ago ahead of the export the modeller just made. Equal times
// keep the list order, and a createdAt that does not parse is the oldest time.
export function latestInboxItem(items: InboxItem[]): InboxItem | null {
  let latest: InboxItem | null = null;
  let latestAt = -Infinity;
  for (const item of items) {
    const at = Date.parse(item.createdAt) || 0;
    if (at > latestAt) {
      latest = item;
      latestAt = at;
    }
  }
  return latest;
}
