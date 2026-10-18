// What a flow puts in the deck and says about it: the record an insert or a
// repaint stages, how a failed flow reads for the record that rode with it, and
// the sentence a result gains. Invariant: a repaint keeps what the deck
// remembered of a link (its label and first-sent time) over what it can derive.

import { buildDeckRecord, type DeckRecord } from "../link/deck-record";
import {
  sourceLabel,
  type InboxItem,
  type LinkKind,
  type Payload,
} from "../link/model";
import { isDrawTimeout } from "./chart-draw";
import { heldRecordOf } from "./deck-store";
import type { FoundLink } from "./host";
import { isMissingShape } from "./missing-shape";

export type Outcome = "done" | "failed" | "timeout" | "missing";

// The error as the host raised it: a sentence that wraps one keeps it as its
// cause, and must not hide a timeout or a missing shape from the record.
function rootOf(error: unknown): unknown {
  let at = error;
  for (let depth = 0; depth < 5; depth += 1) {
    if (!(at instanceof Error) || at.cause === undefined) break;
    at = at.cause;
  }
  return at;
}

// What a flow's error says about the record that rode with it: a host that
// stopped answering says nothing, a shape gone says nothing about the record.
export function outcomeOf(error: unknown): Outcome {
  const root = rootOf(error);
  if (isDrawTimeout(root)) return "timeout";
  return isMissingShape(root) ? "missing" : "failed";
}

export function joinNotes(
  note: string | undefined,
  more: string | undefined,
): string | undefined {
  if (note === undefined || note === "") return more;
  return more === undefined ? note : `${note}. ${more}`;
}

const now = (): string => new Date().toISOString();

// What an insert puts in the deck: the Inbox row's own words about the link.
export function insertRecord(
  item: InboxItem,
  payload: Payload,
  rev: number,
): DeckRecord {
  return buildDeckRecord(
    {
      id: item.id,
      kind: item.kind,
      label: item.label,
      ...(item.project ? { project: item.project } : {}),
      token: item.token,
      sentAt: item.createdAt,
    },
    payload,
    rev,
    now(),
  );
}

// What a repaint puts in the deck: what the deck remembered of the link, else
// what the source says it is. A deck with only tags has no first-sent time, so
// the push that earns it its record stands in.
export function refreshRecord(
  found: FoundLink,
  payload: Payload,
  rev: number,
  kind: LinkKind = found.tag.kind,
): DeckRecord {
  const prior = heldRecordOf(found.tag.id);
  // A link that changed kind (a picture that became a table) is named anew.
  const label = kind === found.tag.kind ? prior?.label : undefined;
  return buildDeckRecord(
    {
      id: found.tag.id,
      kind,
      label: label ?? sourceLabel(payload.src, kind),
      ...(found.tag.project ? { project: found.tag.project } : {}),
      token: found.token,
      sentAt: prior?.sentAt ?? payload.pushedAt,
    },
    payload,
    rev,
    now(),
  );
}
