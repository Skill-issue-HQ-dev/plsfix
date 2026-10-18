// A deck from before the deck store holds tags and no records, and a link the
// deck agrees with is never repainted: its record is saved from the payload the
// fetch brought back, one batch for the whole deck, with no repaint.
// Invariant: a record is written once per link, and only where the deck has none.

import { sourceLabel, type LinkId } from "../link/model";
import { refreshRecord } from "./deck-build";
import { saveRecords } from "./deck-flow";
import { holdsDeckRecord } from "./deck-store";
import type { RefreshRequest } from "./host";
import { nudgeSave } from "./save-nudge";

export async function saveUnrecorded(
  requests: RefreshRequest[],
  summary: { notes: string[] },
): Promise<void> {
  const byLink = new Map<LinkId, RefreshRequest>();
  for (const one of requests) {
    const id = one.found.tag.id;
    if (!byLink.has(id) && !holdsDeckRecord(id)) byLink.set(id, one);
  }
  if (byLink.size === 0) return;
  const records = [...byLink.values()].map((one) =>
    refreshRecord(one.found, one.payload, one.rev),
  );
  for (const [id, note] of await saveRecords(records)) {
    const { found } = byLink.get(id)!;
    summary.notes.push(`${sourceLabel(found.tag.src, found.tag.kind)} ${note}`);
  }
  // Records alone are no edit on the web: the deck would not be saved.
  await nudgeSave();
}
