// Makes PowerPoint for the web save a deck whose only change is a tag or a
// custom XML part. The web's editor does not count either as an edit: a Break,
// a Remove from this deck or a record written alone is applied to the open deck
// but never saved, and a deck closed right after comes back with the link and
// its key (rig 05.10: the editor said "Saving" for its own reasons and the
// reopened deck still held them; any write after the change flushes them, and
// the edit that costs the user least is the deck's own title written back to the
// value it already holds). Owns that one write. Invariant: it changes nothing a
// person can see, runs on the web only, and never throws: the flow it follows
// has already landed and said so.

import { drawHost, withSyncDeadline } from "./chart-draw";
import { hasPowerPointApi } from "./shapes";

// Document properties arrived in PowerPointApi 1.7, the set the deck store
// itself needs, so a host that keeps parts can always be nudged.
const PROPERTIES_API = "1.7";
const NUDGING = "saving the deck";

// Called by every flow that wrote tags or parts and nothing else, once per
// user action, after its own batch has landed: Break link, Remove from this
// deck, Change source, the record backfill. Two syncs: the title is read, then
// written back as it was.
export async function nudgeSave(): Promise<void> {
  if (drawHost() !== "web" || !hasPowerPointApi(PROPERTIES_API)) return;
  try {
    await PowerPoint.run(async (context) => {
      const properties = context.presentation.properties;
      properties.load("title");
      await withSyncDeadline(context.sync(), NUDGING);
      const title = properties.title;
      properties.title = title;
      await withSyncDeadline(context.sync(), NUDGING);
    });
  } catch {
    // Best effort: a deck the host will not nudge is saved by its next edit.
  }
}
