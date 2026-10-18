// src/ppt/layered-relay.ts
// The RelayApi every deck flow runs against: the active transport (the relay
// client, or this computer's LocalStore) with the deck's own records as the
// floor. The transport is asked first, always, so local mode never reaches the
// relay (design 2026-09-23 s.6) and newer data comes from it alone. Only
// "nothing there" (local notPasted, relay missing) falls to the deck, and only
// for a record that holds the link's data.
// Invariant: the deck answers at the rev the slide shows and never paints a
// byte; a wrong key or a relay fault is the transport's own answer, kept.

import type { DeckRecord } from "../link/deck-record";
import type { LinkId } from "../link/model";
import {
  isRelayError,
  type FetchQuery,
  type FetchResult,
  type InboxRow,
  type OmittedLink,
  type RelayApi,
  type StatusQuery,
  type TouchQuery,
} from "../link/relay";
import type { RelayStatus } from "../link/status";
import { stampSeconds } from "../ui/time";

export type DeckRecords = () => ReadonlyMap<LinkId, DeckRecord>;

// The record that can stand in for a link: one that holds its data. A record
// of identity only (a picture, a link too big for the deck's caps) answers
// nothing, so the row keeps the words the transport earned.
function floorOf(
  held: ReadonlyMap<LinkId, DeckRecord>,
  id: LinkId,
): DeckRecord | undefined {
  const record = held.get(id);
  return record?.dataKept === true && record.payload !== null
    ? record
    : undefined;
}

// A fetch names the rev the slide shows. The deck vouches for exactly that
// one: a record at another rev is not the slide's state, and nothing newer is
// ever taken from it.
function holdsRev(
  held: ReadonlyMap<LinkId, DeckRecord>,
  id: LinkId,
  rev: number | undefined,
): boolean {
  const record = floorOf(held, id);
  return record !== undefined && rev !== undefined && record.rev === rev;
}

export class LayeredRelay implements RelayApi {
  constructor(
    private readonly active: RelayApi,
    private readonly records: DeckRecords,
  ) {}

  async status(items: StatusQuery[]): Promise<RelayStatus[]> {
    const answers = await this.active.status(items);
    const held = this.records();
    return answers.map((answer) => {
      if (answer.rev !== null || answer.error !== undefined) return answer;
      const record = floorOf(held, answer.id);
      if (record === undefined) return answer;
      return {
        id: answer.id,
        rev: record.rev,
        pushedAt: stampSeconds(record.pushedAt),
        deck: true,
        ...(answer.local === true ? { local: true as const } : {}),
      };
    });
  }

  async fetchLinks(items: FetchQuery[]): Promise<FetchResult> {
    const answer = await this.active.fetchLinks(items);
    const held = this.records();
    const known = new Map(items.map((item) => [item.id, item.knownRev]));
    const omitted = answer.omitted.map((entry): OmittedLink =>
      (entry.reason === "missing" || entry.reason === "notPasted") &&
      holdsRev(held, entry.id, known.get(entry.id))
        ? { id: entry.id, reason: "inDeck" }
        : entry,
    );
    return { items: answer.items, omitted };
  }

  async getLink(
    id: string,
    auth: string,
    knownRev?: number,
  ): Promise<{ rev: number; blob: Uint8Array } | "unchanged"> {
    try {
      return await this.active.getLink(id, auth, knownRev);
    } catch (error) {
      const gone = isRelayError(error) && error.kind === "missing";
      if (gone && holdsRev(this.records(), id, knownRev)) return "unchanged";
      throw error;
    }
  }

  // The deck keeps only the version on the slide: no floor under a revert.
  getLinkRev(
    id: string,
    auth: string,
    rev: number,
  ): Promise<{ rev: number; blob: Uint8Array }> {
    return this.active.getLinkRev(id, auth, rev);
  }

  putLink(
    id: string,
    auth: string,
    blob: Uint8Array,
    currentRev?: number,
  ): Promise<{ rev: number }> {
    return this.active.putLink(id, auth, blob, currentRev);
  }

  deleteLink(id: string, auth: string): Promise<void> {
    return this.active.deleteLink(id, auth);
  }

  touchLinks(items: TouchQuery[]): Promise<number> {
    return this.active.touchLinks(items);
  }

  postInbox(
    ws: string,
    auth: string,
    id: string,
    blob: Uint8Array,
  ): Promise<void> {
    return this.active.postInbox(ws, auth, id, blob);
  }

  listInbox(ws: string, auth: string): Promise<InboxRow[]> {
    return this.active.listInbox(ws, auth);
  }

  deleteInbox(ws: string, auth: string, id: string): Promise<void> {
    return this.active.deleteInbox(ws, auth, id);
  }
}

export function layeredRelay(active: RelayApi, records: DeckRecords): RelayApi {
  return new LayeredRelay(active, records);
}
