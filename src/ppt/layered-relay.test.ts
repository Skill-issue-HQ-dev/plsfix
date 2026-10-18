// src/ppt/layered-relay.test.ts
// The floor under the transport, rule by rule: the transport answers first and
// its answer stands; the deck answers only "nothing there", only from a record
// that holds the data, and only at the rev the slide shows.

import { describe, expect, it, vi } from "vitest";
import {
  buildDeckRecord,
  identityOnly,
  type DeckRecord,
} from "../link/deck-record";
import { RelayError } from "../link/relay";
import { FakeRelay } from "../../test/fakerelay";
import { layeredRelay } from "./layered-relay";

const ID = "0123456789abcdef0123456789abcdef";
const OTHER = "fedcba9876543210fedcba9876543210";
const SENT = "2026-10-02T14:02:00.000Z";

function record(id: string, rev: number): DeckRecord {
  return buildDeckRecord(
    { id, kind: "text", label: "Model!B4:F12 text", token: "t", sentAt: SENT },
    {
      v: 1,
      kind: "text",
      text: "EUR 3.2m",
      src: { workbook: "M.xlsx", sheet: "Model", ref: "B4:F12", anchor: "A" },
      pushedAt: SENT,
      hash: "0".repeat(64),
    },
    rev,
    SENT,
  );
}

function floorOf(...records: DeckRecord[]): () => Map<string, DeckRecord> {
  return () => new Map(records.map((one) => [one.id, one]));
}

const blob = new Uint8Array([1]);

describe("status", () => {
  it("passes what the transport holds straight through, the deck's record unread", async () => {
    const relay = new FakeRelay();
    const { rev } = await relay.putLink(ID, "a", blob);
    const layered = layeredRelay(relay, floorOf(record(ID, rev + 7)));
    const [answer] = await layered.status([{ id: ID, auth: "a" }]);
    expect(answer).toEqual({ id: ID, rev, pushedAt: relay.now });
  });

  it("answers from the record when the relay has nothing, at the record's rev and send time", async () => {
    const layered = layeredRelay(new FakeRelay(), floorOf(record(ID, 3)));
    const answers = await layered.status([
      { id: ID, auth: "a" },
      { id: OTHER, auth: "a" },
    ]);
    expect(answers).toEqual([
      { id: ID, rev: 3, pushedAt: Date.parse(SENT) / 1000, deck: true },
      { id: OTHER, rev: null, pushedAt: null },
    ]);
  });

  it("keeps the local flag the transport set, so a record at another rev falls back to its words", async () => {
    const transport = new FakeRelay();
    vi.spyOn(transport, "status").mockResolvedValue([
      { id: ID, rev: null, pushedAt: null, local: true },
    ]);
    const layered = layeredRelay(transport, floorOf(record(ID, 3)));
    expect(await layered.status([{ id: ID, auth: "a" }])).toEqual([
      {
        id: ID,
        rev: 3,
        pushedAt: Date.parse(SENT) / 1000,
        deck: true,
        local: true,
      },
    ]);
  });

  it("does not answer for a record that kept the identity only", async () => {
    const layered = layeredRelay(
      new FakeRelay(),
      floorOf(identityOnly(record(ID, 3))),
    );
    expect(await layered.status([{ id: ID, auth: "a" }])).toEqual([
      { id: ID, rev: null, pushedAt: null },
    ]);
  });

  it("keeps a wrong key as the relay said it", async () => {
    const relay = new FakeRelay();
    await relay.putLink(ID, "owner", blob);
    const layered = layeredRelay(relay, floorOf(record(ID, 1)));
    expect(await layered.status([{ id: ID, auth: "thief" }])).toEqual([
      { id: ID, rev: null, pushedAt: null, error: "auth" },
    ]);
  });

  it("lets a transport's own failure through untouched", async () => {
    const relay = new FakeRelay();
    const fault = new RelayError("server", "relay: 503", 503);
    vi.spyOn(relay, "status").mockRejectedValue(fault);
    await expect(
      layeredRelay(relay, floorOf(record(ID, 1))).status([
        { id: ID, auth: "a" },
      ]),
    ).rejects.toBe(fault);
  });
});

describe("fetchLinks", () => {
  it("says inDeck for a link the transport lacks when the slide shows the record's rev", async () => {
    const layered = layeredRelay(new FakeRelay(), floorOf(record(ID, 3)));
    expect(
      await layered.fetchLinks([{ id: ID, auth: "a", knownRev: 3 }]),
    ).toEqual({ items: [], omitted: [{ id: ID, reason: "inDeck" }] });
  });

  it("hands nothing back for any other rev, or none asked: the deck never paints", async () => {
    const layered = layeredRelay(new FakeRelay(), floorOf(record(ID, 3)));
    for (const knownRev of [2, 4, undefined]) {
      const query = knownRev === undefined ? {} : { knownRev };
      const answer = await layered.fetchLinks([
        { id: ID, auth: "a", ...query },
      ]);
      expect(answer.items).toEqual([]);
      expect(answer.omitted).toEqual([{ id: ID, reason: "missing" }]);
    }
  });

  it("leaves every other answer the transport gave as it was", async () => {
    const relay = new FakeRelay();
    const { rev } = await relay.putLink(ID, "a", blob);
    await relay.putLink(OTHER, "owner", blob);
    const layered = layeredRelay(
      relay,
      floorOf(record(ID, rev), record(OTHER, 1)),
    );
    const answer = await layered.fetchLinks([
      { id: ID, auth: "a", knownRev: rev },
      { id: OTHER, auth: "thief", knownRev: 1 },
    ]);
    expect(answer.omitted).toEqual([
      { id: ID, reason: "unchanged" },
      { id: OTHER, reason: "auth" },
    ]);
  });
});

describe("getLink", () => {
  it("answers unchanged for a link the relay forgot when the slide shows the record's rev", async () => {
    const layered = layeredRelay(new FakeRelay(), floorOf(record(ID, 3)));
    expect(await layered.getLink(ID, "a", 3)).toBe("unchanged");
  });

  it("keeps the relay's missing for any other rev and for no rev asked", async () => {
    const layered = layeredRelay(new FakeRelay(), floorOf(record(ID, 3)));
    for (const knownRev of [2, undefined]) {
      await expect(layered.getLink(ID, "a", knownRev)).rejects.toMatchObject({
        kind: "missing",
      });
    }
  });

  it("keeps a relay fault as it was said", async () => {
    const relay = new FakeRelay();
    const fault = new RelayError("server", "relay: 503", 503);
    vi.spyOn(relay, "getLink").mockRejectedValue(fault);
    await expect(
      layeredRelay(relay, floorOf(record(ID, 3))).getLink(ID, "a", 3),
    ).rejects.toBe(fault);
  });
});

describe("everything else", () => {
  it("is the transport's: no floor under a revert, and the Inbox is the transport's own", async () => {
    const relay = new FakeRelay();
    const layered = layeredRelay(relay, floorOf(record(ID, 3)));
    await expect(layered.getLinkRev(ID, "a", 2)).rejects.toMatchObject({
      kind: "missing",
    });
    await layered.postInbox("ws", "a", ID, blob);
    expect((await layered.listInbox("ws", "a")).map((row) => row.id)).toEqual([
      ID,
    ]);
    await layered.deleteInbox("ws", "a", ID);
    expect(await layered.listInbox("ws", "a")).toEqual([]);
  });
});
