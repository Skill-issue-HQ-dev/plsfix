// How a link flow carries the deck's copy of its object: the record staged, its
// ride in the batch that tags the object (no round trip of its own), its save
// in a batch of its own when it cannot ride, and the hooks host.ts wears.
// Invariant: a record never fails a flow the host would have run without it.

import type { DeckRecord } from "../link/deck-record";
import type { InboxItem, LinkId, Payload } from "../link/model";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import { loadDeckStore } from "./deck-boot";
import { insertRecord, joinNotes, outcomeOf, type Outcome } from "./deck-build";
import {
  adoptDeckWrite,
  deckStoreAvailable,
  deckStoreStale,
  deleteDeckRecord,
  distrustDeckStore,
  holdsDeckRecord,
  notSavedNote,
  releaseDeckWrite,
  writeDeckRecord,
  type DeckWrite,
} from "./deck-store";
import type { FoundLink, InsertResult } from "./host";
import type { InsertTarget } from "./placement";
import type { ScanSeen } from "./scan-settle";

// A write the host refuses is asked once more after the deck is read again: a
// part id the index held may be gone.
const ATTEMPTS = 2;
const SAVING = "saving the deck's copy of the links";
const REMOVING = "removing the deck's copy of the link";

// A record waiting for the flow that writes its object to finish: the write
// that rode in that flow's batch, once it did, or why the budget refused it.
interface Staged {
  record: DeckRecord;
  write?: DeckWrite;
  refusal?: string;
}

const staged = new Map<LinkId, Staged>();
// A table repaints cell by cell and knows its payload, not its link: the
// payload object the flow was handed leads back to the record staged for it.
const idOfPayload = new WeakMap<object, LinkId>();
// How many objects of each link the last scan, insert or break left in the
// deck: a record goes with the last of them, never with the first.
const copies = new Map<LinkId, number>();
// Off for the rest of the session once the host refuses a batch that carried a
// record: records then go in on their own, in a batch after the flow's.
let riding = true;
// A Break in flight, by link: the record's delete rides the batch that drops
// the tags when this object is the last the deck holds.
const breaking = new Map<LinkId, { rode: boolean }>();

export function deckRiding(): boolean {
  return riding;
}

// True when this call took the record on: a flow inside another one that
// already staged it leaves the finishing to its owner.
export function stageDeckRecord(record: DeckRecord, payload: object): boolean {
  if (!deckStoreAvailable() || staged.has(record.id)) return false;
  staged.set(record.id, { record });
  idOfPayload.set(payload, record.id);
  return true;
}

// Called inside the batch that changes the object's tags, before its sync: the
// record is queued there (written, or deleted when the batch is a Break's) and
// costs no round trip. It stays out of a batch when the deck has stopped
// riding, or its index is in doubt - finish saves it then.
export function rideDeckCopy(
  context: PowerPoint.RequestContext,
  id: LinkId,
): void {
  const flight = breaking.get(id);
  if (flight !== undefined) {
    rideRemoval(context, id, flight);
    return;
  }
  const entry = staged.get(id);
  if (entry === undefined || entry.write || entry.refusal) return;
  if (!riding || deckStoreStale()) return;
  const result = writeDeckRecord(context, entry.record);
  if (result.ok) entry.write = result.write;
  else entry.refusal = result.note;
}

function rideRemoval(
  context: PowerPoint.RequestContext,
  id: LinkId,
  flight: { rode: boolean },
): void {
  if (!riding || deckStoreStale() || (copies.get(id) ?? 1) > 1) return;
  flight.rode ||= deleteDeckRecord(context, id);
}

export function rideDeckCopyOf(
  context: PowerPoint.RequestContext,
  payload: object,
): void {
  const id = idOfPayload.get(payload);
  if (id !== undefined) rideDeckCopy(context, id);
}

export interface DeckEntry {
  record: DeckRecord;
  payload: object;
}

// Runs one flow that writes the objects of `entries` and sees their records
// into the deck as its outcome allows: the notes are what the flow could not
// keep. A flow the host refuses with a record in its batch fails as it always
// did (an insert is not safe to repeat; links.ts paintBatch runs a picture
// batch again row by row) unless its caller says `again`: a repaint writes the
// same however often it runs, so it is run once more with nothing riding, and
// the records are saved on their own after the deck is read again.
export async function withDeckRecords<T>(
  entries: DeckEntry[],
  run: () => Promise<T>,
  again = false,
): Promise<{ result: T; notes: Map<LinkId, string> }> {
  const owned = entries
    .filter((entry) => stageDeckRecord(entry.record, entry.payload))
    .map((entry) => entry.record.id);
  let result: T;
  try {
    result = await run();
  } catch (error) {
    if (!again || !ridden(owned) || outcomeOf(error) === "timeout") {
      await finishDeckRecords(owned, outcomeOf(error));
      throw error;
    }
    try {
      result = await run();
    } catch (second) {
      await finishDeckRecords(owned, outcomeOf(second));
      throw second;
    }
  }
  return { result, notes: await finishDeckRecords(owned, "done") };
}

// True when a record rode the batch that has just failed.
function ridden(ids: LinkId[]): boolean {
  return ids.some((id) => staged.get(id)?.write !== undefined);
}

// What is left of the staged records once their flow has ended. A write that
// rode and landed is the deck's; one that rode and did not is why the index
// is in doubt (and, after a failure, why the deck stops riding); a record that
// never rode is saved in a batch of its own when the flow succeeded. Never
// throws: the notes are what could not be kept.
export async function finishDeckRecords(
  ids: LinkId[],
  outcome: Outcome,
): Promise<Map<LinkId, string>> {
  const notes = new Map<LinkId, string>();
  const toSave: DeckRecord[] = [];
  for (const id of ids) {
    const entry = staged.get(id);
    if (entry === undefined) continue;
    staged.delete(id);
    if (entry.write !== undefined && adoptDeckWrite(entry.write)) {
      if (entry.write.note !== undefined) notes.set(id, entry.write.note);
      continue;
    }
    if (entry.refusal !== undefined) {
      notes.set(id, entry.refusal);
    } else if (outcome === "done") {
      toSave.push(entry.record);
    }
    if (entry.write === undefined) continue;
    releaseDeckWrite(entry.write);
    distrustDeckStore();
    if (outcome === "failed") riding = false;
  }
  for (const [id, note] of await saveRecords(toSave)) notes.set(id, note);
  return notes;
}

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

// The records in one batch and one sync of their own. A host that refuses it
// is asked once more after the deck's parts are read again - a part id the
// index held may be gone - and then the reason is the note.
export async function saveRecords(
  records: DeckRecord[],
): Promise<Map<LinkId, string>> {
  const notes = new Map<LinkId, string>();
  if (records.length === 0 || !deckStoreAvailable()) return notes;
  let reason = "";
  for (let attempt = 0; attempt < ATTEMPTS; attempt += 1) {
    if (deckStoreStale()) await loadDeckStore();
    notes.clear();
    try {
      const writes = await PowerPoint.run(async (context) => {
        const queued: DeckWrite[] = [];
        for (const record of records) {
          const result = writeDeckRecord(context, record);
          if (result.ok) queued.push(result.write);
          else notes.set(record.id, result.note);
        }
        if (queued.length > 0) await withSyncDeadline(context.sync(), SAVING);
        return queued;
      });
      if (writes.filter((write) => !adoptDeckWrite(write)).length === 0) {
        for (const { record, note } of writes) {
          if (note !== undefined) notes.set(record.id, note);
        }
        return notes;
      }
      reason = "PowerPoint did not confirm it";
    } catch (error) {
      reason = reasonOf(error);
      if (isDrawTimeout(error)) {
        distrustDeckStore();
        break;
      }
    }
    distrustDeckStore();
  }
  notes.clear();
  for (const record of records) notes.set(record.id, notSavedNote(reason));
  return notes;
}

export function noteScan(found: readonly { tag: { id: LinkId } }[]): void {
  copies.clear();
  for (const { tag } of found) {
    copies.set(tag.id, (copies.get(tag.id) ?? 0) + 1);
  }
}

function addCopy(id: LinkId): void {
  copies.set(id, (copies.get(id) ?? 0) + 1);
}

// True when the object that carried the link was the last one the deck holds.
function releaseCopy(id: LinkId): boolean {
  const left = (copies.get(id) ?? 1) - 1;
  if (left > 0) {
    copies.set(id, left);
    return false;
  }
  copies.delete(id);
  return true;
}

// The record goes with the last object that carried the link. A batch of its
// own, asked again after the deck is read again if the host refuses it. True
// when the record is gone; a Break does not fail over one that stays, the row
// that asked for the removal itself does.
export async function removeRecord(id: LinkId): Promise<boolean> {
  if (!deckStoreAvailable()) return true;
  for (let attempt = 0; attempt < ATTEMPTS; attempt += 1) {
    if (deckStoreStale()) await loadDeckStore();
    if (!holdsDeckRecord(id)) return true;
    try {
      await PowerPoint.run(async (context) => {
        if (deleteDeckRecord(context, id)) {
          await withSyncDeadline(context.sync(), REMOVING);
        }
      });
      // A part that was not where the index said: read the deck, then again.
      if (!deckStoreStale()) return true;
    } catch (error) {
      distrustDeckStore();
      if (isDrawTimeout(error)) return false;
    }
  }
  return false;
}

// An object no longer carries this link: its record goes unless another
// object of the deck still carries it.
export async function forgetDeckLink(id: LinkId): Promise<void> {
  if (releaseCopy(id)) await removeRecord(id);
}

// An object carries another link now (Change source, once its repaint has
// landed): the old link's record goes, the new one's came with the repaint.
export async function moveDeckLink(from: LinkId, to: LinkId): Promise<void> {
  if (from === to) return;
  addCopy(to);
  await forgetDeckLink(from);
}

// The hooks host.ts wears, each wrapping the adapter function of its name.
// A scan reports how many slides it read, from its own first sync.
type Scan = (slides: (count: number) => void) => Promise<FoundLink[]>;
type Insert = (
  item: InboxItem,
  payload: Payload,
  rev: number,
  target?: InsertTarget,
) => Promise<InsertResult>;
type Break = (found: FoundLink) => Promise<void>;

// Told what every scan saw - the slides and the links - the moment it returns:
// the pane's settle (deck-panel.ts) is fed by it, with no round trip of its own.
let scanSeen: ((seen: ScanSeen) => void) | undefined;
export function onDeckScan(listener: (seen: ScanSeen) => void): void {
  scanSeen = listener;
}

export function deckScan(scan: Scan): () => Promise<FoundLink[]> {
  return async () => {
    let slides = 0;
    const found = await scan((count) => {
      slides = count;
    });
    noteScan(found);
    scanSeen?.({ slides, ids: new Set(found.map((one) => one.tag.id)) });
    return found;
  };
}

export function deckInsert(insert: Insert): Insert {
  return async (item, payload, rev, target) => {
    const entry = { record: insertRecord(item, payload, rev), payload };
    const { result, notes } = await withDeckRecords([entry], () =>
      insert(item, payload, rev, target),
    );
    addCopy(item.id);
    const note = joinNotes(result.note, notes.get(item.id));
    return note === result.note ? result : { ...result, note };
  };
}

// A batch that landed left the record gone with the tags. One the host
// refused with the delete in it (a part a co-author replaced is the usual
// cause) left the index in doubt and the tags on the shape: it is run once
// more without the delete, and the record is then found by a fresh read.
export function deckBreak(breakLink: Break): Break {
  return async (found) => {
    const id = found.tag.id;
    const flight = { rode: false };
    breaking.set(id, flight);
    try {
      await breakWithRetry(breakLink, found, flight);
    } finally {
      breaking.delete(id);
    }
    if (!flight.rode) return forgetDeckLink(id);
    copies.delete(id);
    if (deckStoreStale()) await removeRecord(id);
  };
}

async function breakWithRetry(
  breakLink: Break,
  found: FoundLink,
  flight: { rode: boolean },
): Promise<void> {
  try {
    await breakLink(found);
  } catch (error) {
    if (!flight.rode || outcomeOf(error) === "timeout") throw error;
    distrustDeckStore();
    flight.rode = false;
    await breakLink(found);
  }
}
