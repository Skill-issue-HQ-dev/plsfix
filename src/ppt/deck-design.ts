// The deck's own design part, on the host: the read at boot, the write queued
// into a caller's batch or saved on its own, and the part index in between.
// One presentation-level custom XML part in urn:plsfix:deck-design:1 (the codec
// is deck-design-record.ts); the link store's purge and budget (deck-store.ts)
// never see it. Gated on PowerPointApi 1.7 like that store.
// Invariants: a write never throws into a flow - it says why it did not happen -
// and a part from a newer pls,fix is left exactly as it is; a copy that was never
// read, or whose read was refused, is UNKNOWN, not empty: nothing is built from it
// and no part is added until a read has told what the deck really holds.

import {
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  DESIGN_RECORD_MAX_BYTES,
  encodeDeckDesign,
  NEWER_DECK_DESIGN,
  type DeckDesign,
} from "./deck-design-record";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import { deckStoreAvailable } from "./deck-store";
import { nudgeSave } from "./save-nudge";

const READING = "reading the deck's brand colors";
const SAVING = "saving the deck's brand colors";
const ATTEMPTS = 2;
const UNAVAILABLE = "this PowerPoint cannot keep data in a deck";
const TOO_BIG = "the deck has too many slide masters to keep";
export const DESIGN_NOT_SAVED = "The deck's brand colors were not saved";
export const DESIGN_UNREAD = "This deck's saved brand colors could not be read";
const UNREAD_WRITE = "pls,fix could not read what this deck already holds";

/** What the boot read found. */
export interface DeckDesignBoot {
  /** False on a host that predates custom XML parts: nothing was asked of it. */
  supported: boolean;
  /** The design the deck holds, null when it holds none this build can read. */
  design: DeckDesign | null;
  /** Why a part was left alone: one from a newer pls,fix says so. */
  issue?: string;
  /** The host's own words when it refused the read: the deck is then UNKNOWN. */
  failure?: string;
  /** The read never answered within its deadline. */
  hung?: true;
}

export type DesignSaved = { ok: true } | { ok: false; note: string };

/** The design to write, built from what the deck holds NOW (null: nothing to write). */
export type DesignBuilder = (current: DeckDesign | null) => DeckDesign | null;

// The part a design lives in and the design it held when last read or written.
// `spare` are further parts of the namespace this build reads: a second pane's.
// `stale` = the index may not be true: nothing read yet, a read the host refused,
// or a batch that carried a write was given up. Every writer reads again first.
const state = {
  part: null as { id: string; design: DeckDesign } | null,
  spare: [] as string[],
  readOnly: false,
  issue: undefined as string | undefined,
  stale: true,
};

export function designAvailable(): boolean {
  return deckStoreAvailable();
}

/**
 * The design the deck held as of the last read and every write since. While
 * `deckDesignUnknown()` is true this is only the last copy that was known (null at
 * boot), never "none": a caller that builds on it settles the copy first.
 */
export function deckDesign(): DeckDesign | null {
  return state.part?.design ?? null;
}

/** True when the deck may hold a design this copy has not seen. */
export function deckDesignUnknown(): boolean {
  return designAvailable() && state.stale;
}

/** Why the deck's design is read-only, when it is: a newer pls,fix wrote it. */
export function deckDesignIssue(): string | undefined {
  return state.issue;
}

function reasonOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function notSavedNote(why: string): string {
  return `${DESIGN_NOT_SAVED}: ${why}`;
}

/** The sentence for a deck that could not be read, ending in a full stop. */
export function unreadNote(why: string): string {
  return `${DESIGN_UNREAD}: ${why.replace(/[.\s]+$/u, "")}.`;
}

interface ReadPart {
  id: string;
  text: { value: string };
}

// Decodes every part of the namespace: the first readable one is the deck's
// design, later ones are spares, a part from a newer pls,fix freezes writing.
function adoptRead(found: ReadPart[]): void {
  let first: { id: string; design: DeckDesign } | null = null;
  const spare: string[] = [];
  state.readOnly = false;
  state.issue = undefined;
  for (const { id, text } of found) {
    const decoded = decodeDeckDesign(text.value);
    if (decoded.ok) {
      if (first === null) first = { id, design: decoded.design };
      else spare.push(id);
    } else if (decoded.reason === "newerVersion") {
      state.readOnly = true;
      state.issue = decoded.message;
    }
  }
  state.part = first;
  state.spare = spare;
  state.stale = false;
}

// Two syncs - the part ids, then the text of every part - and one when the
// deck holds none.
async function readParts(context: PowerPoint.RequestContext): Promise<void> {
  const parts =
    context.presentation.customXmlParts.getByNamespace(DESIGN_NAMESPACE);
  parts.load("items/id");
  await withSyncDeadline(context.sync(), READING);
  const texts = parts.items.map((part) => ({
    id: part.id,
    text: part.getXml(),
  }));
  if (texts.length > 0) await withSyncDeadline(context.sync(), READING);
  adoptRead(texts);
}

/**
 * Reads the deck's design once, at boot, and again after a write left the index in
 * doubt. A read the host refuses leaves the deck UNKNOWN (`failure`), never empty.
 */
export async function loadDeckDesign(): Promise<DeckDesignBoot> {
  if (!designAvailable()) return { supported: false, design: null };
  try {
    await PowerPoint.run(readParts);
  } catch (error) {
    state.stale = true;
    return {
      supported: true,
      design: null,
      failure: reasonOf(error),
      ...(isDrawTimeout(error) ? { hung: true as const } : {}),
    };
  }
  const issue = state.issue;
  return {
    supported: true,
    design: deckDesign(),
    ...(issue === undefined ? {} : { issue }),
  };
}

type Settled = { ok: true } | { ok: false; why: string; hung: boolean };

// A read when the copy is in doubt, none when it can be trusted.
async function settle(): Promise<Settled> {
  if (!designAvailable() || !state.stale) return { ok: true };
  const read = await loadDeckDesign();
  if (read.failure === undefined) return { ok: true };
  return { ok: false, why: read.failure, hung: read.hung === true };
}

/**
 * Reads the deck's design again when the copy is in doubt (never read, a read the
 * host refused, a batch that carried a write given up). Null when the copy can be
 * trusted afterwards; else the sentence saying why not, and the copy stays unknown.
 * Every flow that builds on what the deck holds starts here.
 */
export async function settleDeckDesign(): Promise<string | null> {
  const settled = await settle();
  return settled.ok ? null : unreadNote(settled.why);
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

/** A write queued in a caller's batch, confirmed once that batch has synced. */
export interface DesignWrite {
  design: DeckDesign;
  part: PowerPoint.CustomXmlPart;
}

export type DesignWriteResult =
  { ok: true; write: DesignWrite } | { ok: false; note: string };

/**
 * Queues the design into the caller's batch - a rewrite of the part the deck
 * holds, else a new part, and the spare parts deleted - and takes no sync of its
 * own. Refused with the reason, nothing queued, when the deck's design is a newer
 * pls,fix's, when it would be larger than any record, when the host has no custom
 * XML parts, or when the copy is in doubt: a part is never added on a guess that
 * the deck holds none (the caller settles the copy first).
 */
export function writeDeckDesign(
  context: PowerPoint.RequestContext,
  design: DeckDesign,
): DesignWriteResult {
  try {
    if (!designAvailable())
      return { ok: false, note: notSavedNote(UNAVAILABLE) };
    if (state.readOnly) return { ok: false, note: NEWER_DECK_DESIGN };
    if (state.stale) return { ok: false, note: notSavedNote(UNREAD_WRITE) };
    const xml = encodeDeckDesign(design);
    if (xml.length > DESIGN_RECORD_MAX_BYTES) {
      return { ok: false, note: notSavedNote(TOO_BIG) };
    }
    const parts = context.presentation.customXmlParts;
    const held = state.part;
    const part = held === null ? parts.add(xml) : parts.getItem(held.id);
    if (held !== null) part.setXml(xml);
    part.load("id");
    for (const id of state.spare) parts.getItem(id).delete();
    return { ok: true, write: { design, part } };
  } catch (error) {
    return { ok: false, note: notSavedNote(reasonOf(error)) };
  }
}

/** After the batch that carried the write has synced: the index holds the design. */
export function adoptDeckDesign(write: DesignWrite): boolean {
  let id: string;
  try {
    id = write.part.id;
  } catch {
    return false;
  }
  state.part = { id, design: write.design };
  state.spare = [];
  state.stale = false;
  return true;
}

/** A batch that carried a write was given up: the part index may be wrong now. */
export function distrustDeckDesign(): void {
  state.stale = true;
}

// The one queue every flow that writes the part runs through: a palette commit
// made while Put in the deck theme is running would otherwise write the part
// back without the colors the put has just saved.
let tail: Promise<unknown> = Promise.resolve();

/** Runs `work` once every earlier design write has finished, won or lost. */
export function withDesignLock<T>(work: () => Promise<T>): Promise<T> {
  const run = tail.then(work, work);
  tail = run.then(
    () => undefined,
    () => undefined,
  );
  return run;
}

async function trySave(design: DeckDesign): Promise<DesignWrite | string> {
  const refused: string[] = [];
  const write = await PowerPoint.run(async (context) => {
    const queued = writeDeckDesign(context, design);
    if (!queued.ok) {
      refused.push(queued.note);
      return null;
    }
    await withSyncDeadline(context.sync(), SAVING);
    return queued.write;
  });
  return write ?? refused[0] ?? notSavedNote(UNAVAILABLE);
}

/**
 * Saves a design in a batch of its own: one sync, then the web's editor is nudged
 * to save a part nothing else changed. The design is built by `build` from what
 * the deck holds AFTER this call's own read, so a copy that was never read, or
 * whose read was refused, can never be written over the deck from. A deck that
 * still cannot be read, or a batch the host refuses, is asked again once; the
 * answer says why when it could not, and never throws. A builder that answers
 * null has nothing to write.
 */
export async function saveDeckDesign(
  build: DesignBuilder,
): Promise<DesignSaved> {
  if (!designAvailable()) return { ok: false, note: notSavedNote(UNAVAILABLE) };
  let reason = "";
  for (let attempt = 0; attempt < ATTEMPTS; attempt += 1) {
    const settled = await settle();
    if (!settled.ok) {
      reason = settled.why;
      if (settled.hung) break;
      continue;
    }
    const design = build(deckDesign());
    if (design === null) return { ok: true };
    try {
      const done = await trySave(design);
      if (typeof done === "string") return { ok: false, note: done };
      if (adoptDeckDesign(done)) {
        await nudgeSave();
        return { ok: true };
      }
      reason = "PowerPoint did not confirm it";
    } catch (error) {
      reason = reasonOf(error);
      state.stale = true;
      if (isDrawTimeout(error)) break;
    }
    state.stale = true;
  }
  return { ok: false, note: notSavedNote(reason) };
}
