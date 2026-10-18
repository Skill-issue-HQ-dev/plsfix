// What the deck actions suites share: a fresh pane over the fake deck with the
// actions and the deck boot imported into the same module graph, an export
// inserted and its object deleted by hand, a deck whose record is cut down to
// its identity, and the tag of a shape read back. Owns no assertions.

import { encodeDeckRecord, identityOnly } from "../src/link/deck-record";
import { TAG_LINK, type InboxItem, type LinkTag } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { DEFAULT_TARGET, type InsertTarget } from "../src/ppt/placement";
import type { FakePptHelpers, FakePptShape, FakePresentation } from "./fakeppt";
import type { FakeRelay } from "./fakerelay";
import { removeAllParts, seedParts } from "./ppt.deck-store.support";
import { bootPpt, memoryStore } from "./ppt.support";
import type * as ActionsModule from "../src/ppt/deck-actions";
import type * as BootModule from "../src/ppt/deck-boot";
import type * as LinksModule from "../src/ppt/links";

export interface ActionsPane {
  links: typeof LinksModule;
  actions: typeof ActionsModule;
  boot: typeof BootModule;
  presentation: FakePresentation;
  helpers: FakePptHelpers;
  relay: FakeRelay;
  ws: Workspace;
}

export async function bootActions(): Promise<ActionsPane> {
  const { links, presentation, helpers, relay } = await bootPpt();
  const actions = await import("../src/ppt/deck-actions");
  const boot = await import("../src/ppt/deck-boot");
  const ws = await createWorkspace(memoryStore());
  return { links, actions, boot, presentation, helpers, relay, ws };
}

export function tagOf(shape: FakePptShape): LinkTag {
  return JSON.parse(shape.tags.get(TAG_LINK)!) as LinkTag;
}

// An export inserted, and the object then deleted by hand: the deck keeps the
// link's record and nothing else.
export async function insertedThenDeleted(
  pane: ActionsPane,
  item: InboxItem,
  target: InsertTarget = DEFAULT_TARGET,
): Promise<void> {
  const placed = await pane.links.insertFromInbox(
    item,
    pane.ws,
    pane.relay,
    target,
  );
  pane.presentation.deleteShape(placed.shapeId);
}

// The deck's record of a link, cut down to its identity as a store that was
// full would have left it, and read in again.
export async function keepIdentityOnly(
  pane: ActionsPane,
  item: InboxItem,
): Promise<void> {
  const record = pane.boot.deckRecords().get(item.id);
  if (record === undefined) throw new Error("test: the deck holds no record");
  await removeAllParts();
  await seedParts([encodeDeckRecord(identityOnly(record))]);
  await pane.boot.loadDeckStore();
}
