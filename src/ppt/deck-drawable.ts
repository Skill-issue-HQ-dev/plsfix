// Whether a deck's own copy of a link can be drawn again on THIS host, asked once:
// Put back, its row's button, Redraw and the Inbox's hide set all read it here.
// Invariant: a "no" carries its reason in the kind's own sentence, never "this
// PowerPoint cannot" for a chart type no host draws yet.

import { pictureIssue } from "../link/chart-model";
import { payloadFromRecord, type DeckRecord } from "../link/deck-record";
import type { LinkId, Payload } from "../link/model";
import { declineReasonFor } from "./charts";
import { deckRecords } from "./deck-boot";

// `kept` is whether the deck holds the link's data at all; `issue` is the
// reason as a clause (it follows a label: "<label>: <issue>. <way>") and `way`
// what the user can do instead; `tag` is the row's short word for it.
export type Drawing =
  | { ok: true; payload: Payload }
  | { ok: false; kept: boolean; issue: string; way: string; tag: string };

export type DrawFrom = (record: DeckRecord) => Drawing;

const NO_DATA_ISSUE = "this deck kept no data for it";
const SEND_AGAIN = "Send it again from Excel.";
const AS_A_PICTURE =
  "Send it again from Excel: the Inbox inserts it as a picture.";

// The same refusal as a line of its own, for a link the deck holds no record of.
export function noDataLine(label: string): string {
  return `${label}: ${NO_DATA_ISSUE}. ${SEND_AGAIN}`;
}

export function drawingFrom(record: DeckRecord): Drawing {
  const payload = payloadFromRecord(record);
  if (payload === null) {
    return {
      ok: false,
      kept: false,
      issue: NO_DATA_ISSUE,
      way: SEND_AGAIN,
      tag: "Data not kept in this deck",
    };
  }
  // A chart whose kind this host does not draw as shapes lands as a picture
  // from the Inbox; the record has no picture bytes to put in its place.
  const note =
    payload.kind === "picture" && payload.chart !== undefined
      ? declineReasonFor(payload.chart)
      : null;
  if (note === null) return { ok: true, payload };
  return {
    ok: false,
    kept: true,
    issue: pictureIssue(note),
    way: AS_A_PICTURE,
    tag: "Cannot be drawn here",
  };
}

// The links the Inbox leaves to the list: the ones its Put back can draw. A
// record that cannot be drawn keeps its Inbox row, which is the way back.
export function recordedLinkIds(
  records: ReadonlyMap<LinkId, DeckRecord> = deckRecords(),
): Set<LinkId> {
  return new Set(
    [...records.values()]
      .filter((record) => drawingFrom(record).ok)
      .map((record) => record.id),
  );
}
