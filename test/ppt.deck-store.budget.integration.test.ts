// The deck store's byte budget against the fake deck: a write that would push
// the namespace past 4 MB is refused with the reason, one that fits is not, and
// a rewrite is counted by what it adds. The rest of the store's moves are
// ppt.deck-store.core.integration.test.ts.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  buildDeckRecord,
  DECK_NAMESPACE,
  DECK_STORE_MAX_BYTES,
  encodeDeckRecord,
  identityOnly,
} from "../src/link/deck-record";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import {
  idOf,
  recordOf,
  recordsIn,
  seedParts,
  tablePayload,
  xmlOf,
} from "./ppt.deck-store.support";
import { bootPpt } from "./ppt.support";
import type * as StoreModule from "../src/ppt/deck-store";

enableStrictLoadSemantics();

let store: typeof StoreModule;
let presentation: FakePresentation;

beforeEach(async () => {
  ({ presentation } = await bootPpt());
  store = await import("../src/ppt/deck-store");
});
afterEach(() => {
  uninstallFakePpt();
});

const NS = `xmlns="${DECK_NAMESPACE}"`;

function read(): Promise<unknown> {
  return PowerPoint.run((context) => store.readDeckRecords(context));
}

describe("the byte budget", () => {
  // A part of this namespace the pane cannot read, sized so exactly `room`
  // bytes of the 4 MB are left once it is read.
  const pad = (body: string): string =>
    `<pad ${NS} v="1" id="${idOf(97)}">${body}</pad>`;
  const fillerLeaving = (room: number): string =>
    pad("A".repeat(DECK_STORE_MAX_BYTES - room - pad("").length));

  async function leaving(room: number): Promise<void> {
    await seedParts([fillerLeaving(room)]);
    await read();
    expect(store.deckStoreBytes()).toBe(DECK_STORE_MAX_BYTES - room);
  }

  const table = (rev: number): ReturnType<typeof buildDeckRecord> =>
    buildDeckRecord(
      {
        id: idOf(1),
        kind: "table",
        label: "Model!B4:F12 table",
        token: "token-1",
        sentAt: "2026-10-02T09:00:00.000Z",
      },
      tablePayload(20, 10),
      rev,
      "2026-10-03T08:00:00.000Z",
    );

  it("keeps a record's identity when only its data would push the deck past 4 MB, and says the data was not kept", async () => {
    await leaving(1500);
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, table(1));
      if (!result.ok) throw new Error(result.note);
      expect(result.write.note).toBe(
        "The deck keeps this link's identity but not its data: the link store in this deck is full",
      );
      expect(result.write.record).toMatchObject({
        dataKept: false,
        payload: null,
      });
      await context.sync();
      expect(store.adoptDeckWrite(result.write)).toBe(true);
    });
    expect(recordsIn(presentation)).toMatchObject([
      { id: idOf(1), rev: 1, dataKept: false, payload: null },
    ]);
    expect(store.deckStoreBytes()).toBeLessThanOrEqual(DECK_STORE_MAX_BYTES);
  });

  it("carries that note out of a save in a batch of its own", async () => {
    const flow = await import("../src/ppt/deck-flow");
    await leaving(1500);
    const notes = await flow.saveRecords([table(1)]);
    expect([...notes]).toEqual([
      [
        idOf(1),
        "The deck keeps this link's identity but not its data: the link store in this deck is full",
      ],
    ]);
    expect(recordsIn(presentation)).toMatchObject([
      { id: idOf(1), rev: 1, dataKept: false },
    ]);
  });

  it("says nothing when the whole record fits", async () => {
    await leaving(xmlOf(1).length);
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, recordOf(1));
      if (!result.ok) throw new Error(result.note);
      expect(result.write.note).toBeUndefined();
      await context.sync();
    });
  });

  it("refuses a write that would push the deck past 4 MB even as identity only, and says the store is full", async () => {
    await leaving(100);
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, table(1));
      expect(result.ok).toBe(false);
      expect(!result.ok && result.note).toMatch(
        /^The deck's copy of this link was not saved: the link store in this deck is full$/,
      );
      await context.sync();
    });
    expect(recordsIn(presentation)).toEqual([]);
  });

  it("lets a record that fits through the same deck", async () => {
    await leaving(xmlOf(1).length);
    await PowerPoint.run(async (context) => {
      expect(store.writeDeckRecord(context, recordOf(1)).ok).toBe(true);
      await context.sync();
    });
    expect(recordsIn(presentation)).toHaveLength(1);
  });

  it("refuses a record that is one byte too many when there is no leaner form of it", async () => {
    const lean = identityOnly(recordOf(1));
    await leaving(encodeDeckRecord(lean).length - 1);
    await PowerPoint.run((context) => {
      expect(store.writeDeckRecord(context, lean).ok).toBe(false);
      return Promise.resolve();
    });
  });

  it("falls back to identity only exactly when the identity fits", async () => {
    const whole = recordOf(1, { text: "a".repeat(200) });
    const room = encodeDeckRecord(identityOnly(whole)).length;
    expect(encodeDeckRecord(whole).length).toBeGreaterThan(room);
    await leaving(room);
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, whole);
      expect(result.ok).toBe(true);
      await context.sync();
    });
    expect(recordsIn(presentation)).toMatchObject([
      { id: idOf(1), dataKept: false },
    ]);
  });

  it("refuses a record whose identity form is one byte too many", async () => {
    const whole = recordOf(1, { text: "a".repeat(200) });
    await leaving(encodeDeckRecord(identityOnly(whole)).length - 1);
    await PowerPoint.run((context) => {
      const result = store.writeDeckRecord(context, whole);
      expect(!result.ok && result.note).toMatch(
        /link store in this deck is full/,
      );
      return Promise.resolve();
    });
  });

  it("counts the records already queued in the same batch, so a batch cannot overshoot together", async () => {
    await leaving(xmlOf(1).length + 100);
    await PowerPoint.run(async (context) => {
      const first = store.writeDeckRecord(context, recordOf(1));
      const second = store.writeDeckRecord(context, recordOf(2));
      expect([first.ok, second.ok]).toEqual([true, false]);
      await context.sync();
    });
    expect(recordsIn(presentation)).toHaveLength(1);
  });

  it("gives the room of a write back when its batch never landed", async () => {
    await leaving(xmlOf(1).length + 100);
    await PowerPoint.run((context) => {
      const lost = store.writeDeckRecord(context, recordOf(1));
      if (!lost.ok) throw new Error(lost.note);
      store.releaseDeckWrite(lost.write);
      expect(store.writeDeckRecord(context, recordOf(2)).ok).toBe(true);
      return Promise.resolve();
    });
  });

  it("counts only what a rewrite adds, so a held record can always grow a little", async () => {
    const held = xmlOf(1, { text: "a".repeat(200) }).length;
    await leaving(held + 300);
    await PowerPoint.run(async (context) => {
      const first = store.writeDeckRecord(
        context,
        recordOf(1, { text: "a".repeat(200) }),
      );
      if (!first.ok) throw new Error(first.note);
      await context.sync();
      store.adoptDeckWrite(first.write);
    });
    // 300 bytes of room are left, and this record is bigger than that: it
    // goes in only because it replaces the one that is already counted.
    const next = recordOf(1, { rev: 2, text: "b".repeat(200) });
    expect(xmlOf(1, { rev: 2, text: "b".repeat(200) }).length).toBeGreaterThan(
      300,
    );
    await PowerPoint.run(async (context) => {
      const again = store.writeDeckRecord(context, next);
      expect(again.ok).toBe(true);
      await context.sync();
    });
  });
});
