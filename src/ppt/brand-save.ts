// Where a finished palette edit is saved: the deck's design part and the user's
// default (PowerPoint's own copy, in the JSON shape Excel uses), together, inside
// the design lock so no write reaches the part while Put in the deck theme or
// Restore is running. Answers what did not land as sentences, never throws.
// Invariant: the theme colors the part saved are written back as they were.

import type { KeyStore } from "../link/workspace";
import { writeUserDefault } from "./brand-palette";
import {
  deckDesignIssue,
  designAvailable,
  saveDeckDesign,
  withDesignLock,
} from "./deck-design";
import { buildDeckDesign } from "./deck-design-record";

export interface Saved {
  /** What did not land, in the pane's words. */
  notes: string[];
  /** The brand list the user's default holds now, null when it was not written. */
  user: string[] | null;
}

// A part from a newer pls,fix is read-only: the boot said so once, and every
// edit after it is saved as the user's default alone. The record is built inside
// the save, after its own read, so the theme colors it keeps are the deck's own
// and never those of a copy that was not read.
async function saveDeck(brand: readonly string[]): Promise<string | null> {
  if (!designAvailable() || deckDesignIssue() !== undefined) return null;
  const saved = await saveDeckDesign((current) =>
    buildDeckDesign(brand, current?.themes ?? []),
  );
  return saved.ok ? null : saved.note;
}

async function saveUser(
  keyStore: KeyStore,
  brand: readonly string[],
): Promise<{ note: string | null; user: string[] | null }> {
  try {
    const written = await writeUserDefault(keyStore, brand);
    return { note: null, user: written === null ? null : [...written.brand] };
  } catch (error) {
    const why = error instanceof Error ? error.message : String(error);
    return {
      note: `Your default brand colors were not saved: ${why}`,
      user: null,
    };
  }
}

/** Saves the brand list to both places, in the design lock. */
export function saveBrand(
  keyStore: KeyStore,
  brand: readonly string[],
): Promise<Saved> {
  return withDesignLock(async () => {
    const deckNote = await saveDeck(brand);
    const user = await saveUser(keyStore, brand);
    const notes = [deckNote, user.note].filter(
      (note): note is string => note !== null,
    );
    return { notes, user: user.user };
  });
}
