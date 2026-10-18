// The deck's own copy of its links, one PowerPoint custom XML part each
// (src/link/deck-record.ts): the part index, the read at boot, the write and
// delete queued in the caller's batch, the duplicate rule and the byte budget.
// Invariant: a write never throws into a flow - it says why it did not happen.

import {
  decodeDeckRecord,
  DECK_NAMESPACE,
  DECK_RECORD_MAX_BYTES,
  DECK_STORE_MAX_BYTES,
  identityOnly,
  packDeckRecord,
  type DeckRecord,
} from "../link/deck-record";
import type { LinkId } from "../link/model";
import { withSyncDeadline } from "./chart-draw";
import { hasPowerPointApi } from "./shapes";

// CustomXmlPart arrived in PowerPointApi 1.7: every host that can draw shapes
// (1.8) has it, and a deck opened on one that does not simply keeps no copies.
export const DECK_API = "1.7";
const READING = "reading the deck's link data";
const CLEARING = "clearing the deck's link data";
const DUPLICATES = "removing duplicate link data";

export const NOT_SAVED = "The deck's copy of this link was not saved";
export const FULL_NOTE = `${NOT_SAVED}: the link store in this deck is full`;
export const DATA_NOT_KEPT =
  "The deck keeps this link's identity but not its data: the link store in this deck is full";

export function notSavedNote(why: string): string {
  return `${NOT_SAVED}: ${why}`;
}

// The part a link's record lives in, and the record it held when last read or
// written: what a rewrite addresses, and what the budget is counted from.
interface Held {
  partId: string;
  record: DeckRecord;
  bytes: number;
}

// A part of a link the deck holds twice. Reconcile deletes these; until then
// they count against the budget, and a Break takes them with the record.
interface Spare {
  id: LinkId;
  partId: string;
  bytes: number;
}

interface ReadPart {
  partId: string;
  record: DeckRecord;
  bytes: number;
}

// One pane, one deck: the index lives for the pane's lifetime. `stale` is set
// when a batch that carried a write failed, so a part id the index holds may
// no longer be true; writes then add instead of rewriting until the next read.
const state = {
  held: new Map<LinkId, Held>(),
  spare: [] as Spare[],
  foreignBytes: 0,
  issues: [] as string[],
  stale: false,
};

// What writes queued in a batch that has not synced yet will add, by link:
// counted against the budget at once, so a batch of writes cannot overshoot it
// together. Taken back when the write is adopted or its batch is given up.
const queued = new Map<LinkId, number>();

export function deckStoreAvailable(): boolean {
  return (
    typeof PowerPoint !== "undefined" &&
    typeof Office !== "undefined" &&
    hasPowerPointApi(DECK_API)
  );
}

// The records the deck holds as of the last read and every write since.
export function heldRecords(): Map<LinkId, DeckRecord> {
  return new Map(
    [...state.held].map(([id, held]): [LinkId, DeckRecord] => [
      id,
      held.record,
    ]),
  );
}

export function heldRecordOf(id: LinkId): DeckRecord | undefined {
  return state.held.get(id)?.record;
}

// True when the deck holds a part for the link, the spare ones included.
export function holdsDeckRecord(id: LinkId): boolean {
  return state.held.has(id) || state.spare.some((extra) => extra.id === id);
}

// What the namespace costs the deck now: kept records, their spares and the
// parts of this namespace this build cannot read (never deleted, so counted).
export function deckStoreBytes(): number {
  let total = state.foreignBytes;
  for (const held of state.held.values()) total += held.bytes;
  for (const extra of state.spare) total += extra.bytes;
  return total;
}

// Why parts were left alone: a part from a newer pls,fix says so.
export function deckIssues(): string[] {
  return [...state.issues];
}

export function deckStoreStale(): boolean {
  return state.stale;
}

export function distrustDeckStore(): void {
  state.stale = true;
}

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// The better of two records of one link: the higher revision, then the later
// save. A tie keeps whichever was listed first.
function better(a: DeckRecord, b: DeckRecord): boolean {
  return a.rev > b.rev || (a.rev === b.rev && a.savedAt > b.savedAt);
}

// Two parts for one link happen when two panes - a co-author's, or an old
// session's - each wrote one. The winner is the one kept; the rest are spares.
function reconcileParts(parts: ReadPart[]): {
  winners: Map<LinkId, Held>;
  spare: Spare[];
} {
  const winners = new Map<LinkId, Held>();
  const spare: Spare[] = [];
  const spareOf = (part: Held, id: LinkId): Spare => ({
    id,
    partId: part.partId,
    bytes: part.bytes,
  });
  for (const part of parts) {
    const id = part.record.id;
    const current = winners.get(id);
    if (current === undefined) {
      winners.set(id, part);
    } else if (better(part.record, current.record)) {
      winners.set(id, part);
      spare.push(spareOf(current, id));
    } else {
      spare.push(spareOf(part, id));
    }
  }
  return { winners, spare };
}

// Two syncs - the part ids, then the text of every part - and one when the
// deck holds none. A part this build cannot read stays exactly as it is.
export async function readDeckRecords(
  context: PowerPoint.RequestContext,
): Promise<Map<LinkId, DeckRecord>> {
  const found =
    context.presentation.customXmlParts.getByNamespace(DECK_NAMESPACE);
  found.load("items/id");
  await withSyncDeadline(context.sync(), READING);
  const texts = found.items.map((part) => ({
    partId: part.id,
    text: part.getXml(),
  }));
  if (texts.length > 0) await withSyncDeadline(context.sync(), READING);
  const readable: ReadPart[] = [];
  const issues: string[] = [];
  let foreignBytes = 0;
  for (const { partId, text } of texts) {
    const xml = text.value;
    const decoded = decodeDeckRecord(xml);
    if (decoded.ok) {
      readable.push({ partId, record: decoded.record, bytes: xml.length });
      continue;
    }
    foreignBytes += xml.length;
    if (
      decoded.reason === "newerVersion" &&
      !issues.includes(decoded.message)
    ) {
      issues.push(decoded.message);
    }
  }
  const { winners, spare } = reconcileParts(readable);
  Object.assign(state, { held: winners, spare, foreignBytes, issues });
  state.stale = false;
  queued.clear();
  return heldRecords();
}

// What a queued write needs to be confirmed once its batch syncs: the part
// whose id was loaded in that batch, and the record as it was packed.
export interface DeckWrite {
  record: DeckRecord;
  bytes: number;
  part: PowerPoint.CustomXmlPart;
  replacedPartId: string | undefined;
  // Set when the deck kept less than it was given: the words for the row.
  note?: string;
}

export type DeckWriteResult =
  { ok: true; write: DeckWrite } | { ok: false; note: string };

interface Fit {
  xml: string;
  record: DeckRecord;
  note?: string;
}

// The record as the deck can take it: whole when it fits the 4 MB, else as
// identity only when that fits, else not at all. `room` is what the store has
// left for this link, the part it already holds counted back in.
function fit(record: DeckRecord, room: number): Fit | string {
  const whole = packDeckRecord(record);
  if (whole.xml.length > DECK_RECORD_MAX_BYTES) {
    return notSavedNote("the link is too large to keep");
  }
  if (whole.xml.length <= room) return whole;
  const lean = packDeckRecord(identityOnly(whole.record));
  if (lean.xml.length <= room) {
    return { ...lean, note: DATA_NOT_KEPT };
  }
  return FULL_NOTE;
}

// Queues the record into the caller's batch - a rewrite of the part it holds,
// else a new part - and takes no sync of its own. A record over the record cap
// goes in as identity only, and so does one the deck has no room for whole; one
// it has no room for at all, or a host with no custom XML parts, is refused with
// the reason and nothing queued.
export function writeDeckRecord(
  context: PowerPoint.RequestContext,
  record: DeckRecord,
): DeckWriteResult {
  try {
    const held = state.held.get(record.id);
    const rewrite = held !== undefined && !state.stale ? held : undefined;
    let others = 0;
    for (const [id, extra] of queued) if (id !== record.id) others += extra;
    const spent = deckStoreBytes() + others - (rewrite?.bytes ?? 0);
    const fitted = fit(record, DECK_STORE_MAX_BYTES - spent);
    if (typeof fitted === "string") return { ok: false, note: fitted };
    const { xml, record: kept, note } = fitted;
    // A rewrite of a part that is gone is found out by the sync, not here.
    const parts = context.presentation.customXmlParts;
    const part =
      rewrite === undefined ? parts.add(xml) : parts.getItem(rewrite.partId);
    if (rewrite !== undefined) part.setXml(xml);
    part.load("id");
    queued.set(record.id, Math.max(0, xml.length - (rewrite?.bytes ?? 0)));
    const write: DeckWrite = {
      record: kept,
      bytes: xml.length,
      part,
      replacedPartId: rewrite?.partId,
      ...(note === undefined ? {} : { note }),
    };
    return { ok: true, write };
  } catch (error) {
    return { ok: false, note: notSavedNote(reasonOf(error)) };
  }
}

// After the batch that carried the write has synced: its part answers its id,
// and the index holds the record. False when the batch never landed - the
// part's id was loaded in it, so an id that cannot be read says no.
export function adoptDeckWrite(write: DeckWrite): boolean {
  const id = write.record.id;
  let partId: string;
  try {
    partId = write.part.id;
  } catch {
    return false;
  }
  queued.delete(id);
  const before = state.held.get(id);
  if (before !== undefined && before.partId !== partId) {
    state.spare.push({ id, partId: before.partId, bytes: before.bytes });
  }
  state.held.set(id, { partId, record: write.record, bytes: write.bytes });
  return true;
}

// The batch that carried this write is given up: its room is the budget's again.
export function releaseDeckWrite(write: DeckWrite): void {
  queued.delete(write.record.id);
}

// Queues the delete of the link's part and of any spare part of the same link,
// so no copy of its token is left behind. False when the deck holds none. A
// part that is not where the index put it fails the batch at its sync: whoever
// ran the batch distrusts the index then and reads the deck again.
export function deleteDeckRecord(
  context: PowerPoint.RequestContext,
  id: LinkId,
): boolean {
  const held = state.held.get(id);
  const spares = state.spare.filter((extra) => extra.id === id);
  if (held === undefined && spares.length === 0) return false;
  const parts = context.presentation.customXmlParts;
  const ids = [...(held ? [held.partId] : []), ...spares.map((s) => s.partId)];
  for (const partId of ids) parts.getItem(partId).delete();
  state.held.delete(id);
  state.spare = state.spare.filter((extra) => extra.id !== id);
  return true;
}

// Deletes the spare parts a read found, one sync: the winner of each link is
// the one the deck keeps. How many parts went.
export async function reconcileDuplicates(
  context: PowerPoint.RequestContext,
): Promise<number> {
  const parts = context.presentation.customXmlParts;
  const count = state.spare.length;
  for (const extra of state.spare) parts.getItem(extra.partId).delete();
  if (count > 0) await withSyncDeadline(context.sync(), DUPLICATES);
  state.spare = [];
  return count;
}

// Every part of the namespace, the ones this build cannot read too: what the
// deck keeps of pls,fix links, gone. How many parts went.
export async function purgeDeckStore(
  context: PowerPoint.RequestContext,
): Promise<number> {
  const found =
    context.presentation.customXmlParts.getByNamespace(DECK_NAMESPACE);
  found.load("items/id");
  await withSyncDeadline(context.sync(), CLEARING);
  const listed = found.items;
  for (const part of listed) part.delete();
  const count = listed.length;
  if (count > 0) await withSyncDeadline(context.sync(), CLEARING);
  state.held.clear();
  queued.clear();
  Object.assign(state, {
    spare: [],
    foreignBytes: 0,
    issues: [],
    stale: false,
  });
  return count;
}
