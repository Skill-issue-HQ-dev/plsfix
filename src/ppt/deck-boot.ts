// The deck's records, read once when the pane boots. Owns loadDeckStore (the
// boot step) and deckRecords (what the next slices read the deck's rows from).
// Invariant: a deck, a host or a part that cannot be read leaves the records
// empty and never blocks or fails the pane.

import type { DeckRecord } from "../link/deck-record";
import type { LinkId } from "../link/model";
import {
  deckIssues,
  deckStoreAvailable,
  heldRecords,
  readDeckRecords,
  reconcileDuplicates,
} from "./deck-store";

export interface DeckBoot {
  // False on a host that predates custom XML parts: nothing was asked of it.
  supported: boolean;
  records: number;
  // Why parts were left alone (a part from a newer pls,fix), in the pane's words.
  issues: string[];
  // The host's own words when it refused the read.
  failure?: string;
}

// Reads the deck's records and settles any link two panes both wrote. Also the
// way the store is read again once a write has left its index in doubt.
export async function loadDeckStore(): Promise<DeckBoot> {
  if (!deckStoreAvailable()) {
    return { supported: false, records: 0, issues: [] };
  }
  try {
    await PowerPoint.run(async (context) => {
      await readDeckRecords(context);
      await reconcileDuplicates(context);
    });
  } catch (error) {
    const failure = error instanceof Error ? error.message : String(error);
    return { ...summary(), failure };
  }
  return summary();
}

function summary(): DeckBoot {
  return {
    supported: true,
    records: heldRecords().size,
    issues: deckIssues(),
  };
}

// What the deck holds for each link as of the boot read and every write since:
// the Not on a slide rows, Put back and the Inbox's hide set are read from it.
export function deckRecords(): ReadonlyMap<LinkId, DeckRecord> {
  return heldRecords();
}
