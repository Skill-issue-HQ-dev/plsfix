// What a deck's own copy of a link does without Excel: Put back and Redraw. All
// through insertLink / refreshLink / refreshChart, so the record rides the
// object's own batch. Invariant: a record with no data, or a chart this host
// cannot draw as shapes, is refused in words; no picture is drawn.

import {
  INSERT_CANCELLED,
  INSERT_STOPPED,
  pictureIssue,
  REFRESH_STOPPED,
} from "../link/chart-model";
import type { DeckRecord } from "../link/deck-record";
import {
  sourceLabel,
  type InboxItem,
  type LinkId,
  type Payload,
  type PicturePayload,
} from "../link/model";
import { CalledOff } from "../ui/guard";
import { isDrawTimeout } from "./chart-draw";
import { chartPlan, declineReason, refreshChart } from "./charts";
import { deckRecords, loadDeckStore } from "./deck-boot";
import { drawingFrom, noDataLine } from "./deck-drawable";
import { joinNotes, refreshRecord } from "./deck-build";
import { removeRecord, withDeckRecords } from "./deck-flow";
import {
  goToSlide,
  insertLink,
  refreshLink,
  scanLinks,
  type FoundLink,
  type InsertResult,
} from "./host";
import { failureLine, insertNote, type LinkRow } from "./links";
import { missingShapeError } from "./missing-shape";
import { DEFAULT_TARGET, type InsertTarget } from "./placement";
import { tagFor } from "./refresh";
import { nudgeSave } from "./save-nudge";
import { GROUP_TYPE } from "./shapes";
import { isChosen, NOW_CHART } from "./upgrade";

export const NO_RECORD = "This deck no longer holds a copy of that link.";
const PUT_BACK_STOPPED =
  "PowerPoint stopped answering; nothing was inserted, the row is still listed";
const PUT_BACK_CANCELLED =
  "Cancelled: nothing was inserted, the row is still listed";

const DECK_UNREAD =
  "The deck's link data could not be read, so nothing was changed: ";

// This pane's index of the deck's records is as old as its last read, and a
// co-author's Break may have taken a part since: acting on the old index drew
// the object and wrote the record and its link key back into the file. So every
// action reads the file first, and acts on nothing it could not read.
async function readStore(): Promise<void> {
  const read = await loadDeckStore();
  if (read.failure !== undefined) throw new Error(DECK_UNREAD + read.failure);
}

// The record as the file holds it now; one that is gone is refused in words.
async function recordNow(id: LinkId): Promise<DeckRecord> {
  await readStore();
  const record = deckRecords().get(id);
  if (record === undefined) throw new Error(NO_RECORD);
  return record;
}

// The Inbox row a record stands for: what the insert flow takes to make an
// object of a link, here with the link's own id and token.
function itemOf(record: DeckRecord): InboxItem {
  return {
    id: record.id,
    token: record.token,
    kind: record.kind,
    label: record.label,
    src: record.src,
    createdAt: record.sentAt,
    ...(record.project === undefined ? {} : { project: record.project }),
  };
}

// The payload a record draws from, or the sentence for why it cannot: the
// deck's one answer (deck-drawable.ts), the same one that leaves a row's Put
// back dead, so Put back never draws what its row refuses.
function drawable(record: DeckRecord): Payload {
  const drawing = drawingFrom(record);
  if (!drawing.ok) {
    throw new Error(`${record.label}: ${drawing.issue}. ${drawing.way}`);
  }
  return drawing.payload;
}

// The Inbox's own sentences for a stopped or cancelled draw end on "still in
// the Inbox"; there is no Inbox row here, the row of the list stays.
function inDeckWords(error: unknown): unknown {
  if (error instanceof CalledOff && error.message === INSERT_CANCELLED) {
    return new CalledOff(PUT_BACK_CANCELLED);
  }
  if (error instanceof Error && error.message === INSERT_STOPPED) {
    return new Error(PUT_BACK_STOPPED, { cause: error });
  }
  return error;
}

// A slide PowerPoint for the web has not loaded reads as empty until it is
// shown (S1 spike, section 6): the placement would not see what is on it. So a
// slide the pickers name is brought into view before the insert reads it; the
// insert does the same afterwards, so the view ends where it always did.
async function showSlide(slideId: string): Promise<void> {
  try {
    await goToSlide(slideId);
  } catch {
    // The insert reads the slide itself and says why if it cannot.
  }
}

async function insertFrom(
  record: DeckRecord,
  target: InsertTarget,
): Promise<InsertResult> {
  const payload = drawable(record);
  // A selected shape is what that spot reads, and showing a slide drops it.
  if (target.slideId !== null && target.where !== "selected-shape") {
    await showSlide(target.slideId);
  }
  try {
    return await insertLink(itemOf(record), payload, record.rev, target);
  } catch (error) {
    throw inDeckWords(error);
  }
}

// The list a Put back is pressed in may be out of date - an undo brought the
// object back, or a co-author did - and drawing it then would leave two. So the
// deck is read once more. This scan cannot see a slide PowerPoint has not loaded
// yet, which reads as empty: that is the settle's job (scan-settle.ts), and no
// row is offered before two late scans agree.
async function requireGone(
  record: DeckRecord,
  refusal: (slide: number) => string,
): Promise<void> {
  const here = (await scanLinks()).find((one) => one.tag.id === record.id);
  if (here !== undefined) throw new Error(refusal(here.slideIndex + 1));
}

// The object back on a slide, from the deck's copy: the link's own id, token
// and revision, so Excel's next push refreshes it exactly as before.
export async function putBack(
  id: LinkId,
  target: InsertTarget = DEFAULT_TARGET,
): Promise<string> {
  const record = await recordNow(id);
  await requireGone(
    record,
    (slide) =>
      `${record.label} is already on slide ${String(slide)}. The list was out of date and is read again.`,
  );
  const placed = await insertFrom(record, target);
  return `Put ${record.label} back${insertNote(placed)}`;
}

// The record of a link whose object is gone, taken out of the deck: its part,
// any spare part of the same link, and with them the link's key - by Break's own
// path, asked again after the deck is read again. An object a slide holds again
// is Break link's business; a delete the host refuses twice leaves the row.
export async function removeFromDeck(id: LinkId): Promise<string> {
  const record = await recordNow(id);
  await requireGone(
    record,
    (slide) =>
      `${record.label} is on slide ${String(slide)} again, so it stays in this deck. Tick it and press Break link to let it go.`,
  );
  if (!(await removeRecord(id))) {
    // The index was edited for a delete that did not happen: the file decides.
    await loadDeckStore();
    throw new Error(
      `PowerPoint refused to remove ${record.label} from this deck. The row stays: press Remove again.`,
    );
  }
  await nudgeSave();
  return `Removed ${record.label} from this deck.`;
}

// What Redraw did to the rows it was given, in the order the toast says it.
export interface RedrawSummary {
  redrawn: number;
  // No record, or one that keeps the identity only: nothing to draw from.
  noCopy: number;
  // The object holds a later push than the deck's copy: not drawn over.
  newer: number;
  // A picture the user chose: never turned into shapes, here or on an update.
  chosen: number;
  failed: number;
  failures: string[];
  // One line per row that needs one: why it was left, what could not be kept.
  notes: string[];
}

// A chart drawn again where it sits, whatever it is now - a group with bars
// missing, or a picture the chart once had to be. The new group is built first
// and the old shape goes in the sync that groups it (refreshChart), and the
// record rides that same sync.
async function redrawChart(
  found: FoundLink,
  record: DeckRecord,
  payload: PicturePayload,
): Promise<string | undefined> {
  const label = sourceLabel(found.tag.src, found.tag.kind);
  const plan = chartPlan(payload);
  // A record that passed drawingFrom keeps its chart; this is the type's way
  // of saying so.
  if (plan === null) throw new Error(noDataLine(label));
  // A layout that threw is refused in words too; nothing is drawn over it.
  const refused = declineReason(plan);
  if (refused !== null) {
    throw new Error(
      `${label}: ${pictureIssue(refused)}. The object stays as it is.`,
    );
  }
  const entry = { record: refreshRecord(found, payload, record.rev), payload };
  const tag = tagFor(found.tag, payload, record.rev);
  const { notes } = await withDeckRecords(
    [entry],
    () => refreshChart(found, plan, tag),
    true,
  ).catch((error: unknown) => {
    throw isDrawTimeout(error)
      ? new Error(REFRESH_STOPPED, { cause: error })
      : missingShapeError(`redraw ${label}`, error);
  });
  const note = notes.get(found.tag.id);
  // A picture the chart once had to be says it is a chart now.
  return found.type === GROUP_TYPE ? note : joinNotes(NOW_CHART, note);
}

async function redrawOne(
  found: FoundLink,
  record: DeckRecord | undefined,
  summary: RedrawSummary,
): Promise<void> {
  const label = sourceLabel(found.tag.src, found.tag.kind);
  if (isChosen(found)) {
    summary.chosen += 1;
    summary.notes.push(`${label}: you chose this picture, so it stays one.`);
    return;
  }
  const drawing = record === undefined ? undefined : drawingFrom(record);
  if (
    record === undefined ||
    drawing === undefined ||
    (!drawing.ok && !drawing.kept)
  ) {
    summary.noCopy += 1;
    summary.notes.push(noDataLine(label));
    return;
  }
  // A chart this host cannot draw as shapes stays as it is: no picture here.
  if (!drawing.ok) {
    throw new Error(`${label}: ${drawing.issue}. The object stays as it is.`);
  }
  // The deck's copy is never drawn over a later push: ISO stamps compare as text.
  if (record.pushedAt < found.tag.pushedAt) {
    summary.newer += 1;
    summary.notes.push(
      `${label}: the object on the slide is newer than the deck's copy. Update it from Excel instead.`,
    );
    return;
  }
  const payload = drawing.payload;
  const note =
    payload.kind === "picture"
      ? await redrawChart(found, record, payload)
      : await refreshLink(found, payload, record.rev, true);
  summary.redrawn += 1;
  if (note !== undefined) summary.notes.push(`${label} ${note}`);
}

// The rows' objects drawn again from the deck's records, one after the other:
// one row's failure never stops the rest. Its own chart is drawn before the old
// one is taken down, so a failure leaves the object as it was.
export async function redrawLinks(rows: LinkRow[]): Promise<RedrawSummary> {
  const summary: RedrawSummary = {
    redrawn: 0,
    noCopy: 0,
    newer: 0,
    chosen: 0,
    failed: 0,
    failures: [],
    notes: [],
  };
  await readStore();
  const records = deckRecords();
  for (const { found } of rows) {
    try {
      await redrawOne(found, records.get(found.tag.id), summary);
    } catch (error) {
      summary.failed += 1;
      summary.failures.push(failureLine(found, error));
    }
  }
  return summary;
}

export function summarizeRedraw(summary: RedrawSummary): string {
  const parts = [
    [summary.redrawn, "redrawn"],
    [summary.noCopy, "with no data in the deck"],
    [summary.newer, "newer than the copy"],
    [summary.chosen, "kept as the picture you chose"],
    [summary.failed, "failed"],
  ] as const;
  const text = parts
    .filter(([count]) => count > 0)
    .map(([count, word]) => `${String(count)} ${word}`)
    .join(", ");
  return text || "Nothing to redraw";
}

export function redrawDetails(summary: RedrawSummary): string | undefined {
  const lines = [...summary.failures, ...summary.notes];
  return lines.length > 0 ? lines.join("\n") : undefined;
}
