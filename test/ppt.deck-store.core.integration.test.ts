// The deck store's own moves against the fake deck: the boot read, the rule
// for two parts of one link, a write that adds or rewrites inside the caller's
// batch, the byte budget, the delete and the purge. What the link flows do
// with these is ppt.deck-flow.integration.test.ts and its siblings.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  buildDeckRecord,
  DECK_NAMESPACE,
  DECK_RECORD_MAX_BYTES,
  NEWER_DECK_RECORD,
} from "../src/link/deck-record";
import {
  customXmlCommitted,
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  idOf,
  partsIn,
  recordOf,
  recordsIn,
  removeAllParts,
  seedParts,
  tablePayload,
  xmlOf,
} from "./ppt.deck-store.support";
import { bootPpt } from "./ppt.support";
import type * as StoreModule from "../src/ppt/deck-store";

enableStrictLoadSemantics();

let store: typeof StoreModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  ({ presentation, helpers } = await bootPpt());
  store = await import("../src/ppt/deck-store");
});
afterEach(() => {
  uninstallFakePpt();
});

const NS = `xmlns="${DECK_NAMESPACE}"`;
const FOREIGN = `<note ${NS} v="1" id="${idOf(99)}">QUJD</note>`;
const NEWER = `<link ${NS} v="2" id="${idOf(98)}">QUJD</link>`;

function read(): Promise<Map<string, ReturnType<typeof recordOf>>> {
  return PowerPoint.run((context) => store.readDeckRecords(context));
}

describe("the boot read", () => {
  it("reads every record of the namespace in two syncs and hands back id -> record", async () => {
    await seedParts([xmlOf(1), xmlOf(2, { rev: 3 }), xmlOf(3)]);
    const from = helpers.syncCount();
    const records = await read();
    expect(
      helpers.syncCount() - from,
      "one sync for the part ids, one for every part's text",
    ).toBe(2);
    expect([...records.keys()]).toEqual([idOf(1), idOf(2), idOf(3)]);
    expect(records.get(idOf(2))).toEqual(recordOf(2, { rev: 3 }));
    expect(store.heldRecords()).toEqual(records);
  });

  it("reads an empty deck in one sync and holds nothing", async () => {
    const from = helpers.syncCount();
    expect((await read()).size).toBe(0);
    expect(helpers.syncCount() - from).toBe(1);
    expect(store.deckStoreBytes()).toBe(0);
  });

  it("leaves a part outside the namespace out of it", async () => {
    await PowerPoint.run(async (context) => {
      context.presentation.customXmlParts.add(`<other xmlns="urn:other"/>`);
      await context.sync();
    });
    await seedParts([xmlOf(1)]);
    expect([...(await read()).keys()]).toEqual([idOf(1)]);
  });

  it("never touches a part it cannot read, and says what a newer one is", async () => {
    await seedParts([xmlOf(1), FOREIGN, NEWER]);
    expect([...(await read()).keys()]).toEqual([idOf(1)]);
    expect(store.deckIssues()).toEqual([NEWER_DECK_RECORD]);
    expect(store.deckStoreBytes()).toBe(
      partsIn(presentation).reduce((sum, xml) => sum + xml.length, 0),
    );
    const deleted = await PowerPoint.run((context) =>
      store.reconcileDuplicates(context),
    );
    expect(deleted).toBe(0);
    expect(partsIn(presentation)).toEqual([xmlOf(1), FOREIGN, NEWER]);
  });
});

describe("two parts for one link", () => {
  it("keeps the highest revision and deletes the other on reconcile", async () => {
    await seedParts([
      xmlOf(1, { rev: 2 }),
      xmlOf(1, { rev: 5, text: "the newest" }),
      xmlOf(1, { rev: 3 }),
      xmlOf(2),
    ]);
    const records = await read();
    expect(records.get(idOf(1))).toMatchObject({ rev: 5 });
    const deleted = await PowerPoint.run((context) =>
      store.reconcileDuplicates(context),
    );
    expect(deleted).toBe(2);
    expect(recordsIn(presentation).map((one) => [one.id, one.rev])).toEqual([
      [idOf(1), 5],
      [idOf(2), 1],
    ]);
  });

  it("breaks a tie on the revision by the later save, then by the first listed", async () => {
    await seedParts([
      xmlOf(1, { rev: 4, text: "early", savedAt: "2026-10-03T08:00:00.000Z" }),
      xmlOf(1, { rev: 4, text: "late", savedAt: "2026-10-03T09:00:00.000Z" }),
      xmlOf(2, { rev: 4, text: "first" }),
      xmlOf(2, { rev: 4, text: "second" }),
    ]);
    const records = await read();
    const text = (id: string): unknown =>
      (records.get(id)?.payload as { text?: string } | null)?.text;
    expect(text(idOf(1))).toBe("late");
    expect(text(idOf(2))).toBe("first");
  });

  it("counts the part it will delete against the budget until it is gone", async () => {
    await seedParts([xmlOf(1, { rev: 1 }), xmlOf(1, { rev: 2 })]);
    await read();
    const both = partsIn(presentation).reduce(
      (sum, xml) => sum + xml.length,
      0,
    );
    expect(store.deckStoreBytes()).toBe(both);
    await PowerPoint.run((context) => store.reconcileDuplicates(context));
    expect(store.deckStoreBytes()).toBe(xmlOf(1, { rev: 2 }).length);
  });
});

describe("writing a record", () => {
  it("adds a part for a new link in the caller's batch, which the host holds once it syncs", async () => {
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, recordOf(1));
      if (!result.ok) throw new Error(result.note);
      expect(recordsIn(presentation)).toEqual([recordOf(1)]);
      expect(customXmlCommitted(presentation)).toEqual([]);
      await context.sync();
      expect(recordsIn(presentation, true)).toEqual([recordOf(1)]);
      expect(store.adoptDeckWrite(result.write)).toBe(true);
    });
    expect(store.heldRecords()).toEqual(new Map([[idOf(1), recordOf(1)]]));
  });

  it("rewrites that same part the next time, leaving one part and the new revision", async () => {
    await PowerPoint.run(async (context) => {
      const first = store.writeDeckRecord(context, recordOf(1));
      if (!first.ok) throw new Error(first.note);
      await context.sync();
      store.adoptDeckWrite(first.write);
    });
    await PowerPoint.run(async (context) => {
      const next = store.writeDeckRecord(context, recordOf(1, { rev: 2 }));
      if (!next.ok) throw new Error(next.note);
      expect(next.write.replacedPartId).toBeDefined();
      await context.sync();
      expect(store.adoptDeckWrite(next.write)).toBe(true);
    });
    expect(recordsIn(presentation)).toEqual([recordOf(1, { rev: 2 })]);
    expect(store.heldRecords().get(idOf(1))).toMatchObject({ rev: 2 });
  });

  it("rewrites the part the boot read found, not a second one", async () => {
    await seedParts([xmlOf(1), xmlOf(2)]);
    await read();
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, recordOf(2, { rev: 9 }));
      if (!result.ok) throw new Error(result.note);
      await context.sync();
      store.adoptDeckWrite(result.write);
    });
    expect(recordsIn(presentation).map((one) => [one.id, one.rev])).toEqual([
      [idOf(1), 1],
      [idOf(2), 9],
    ]);
  });

  it("queues a rewrite of a part that is gone, which the host refuses at the sync", async () => {
    await seedParts([xmlOf(1)]);
    await read();
    await removeAllParts();
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, recordOf(1, { rev: 2 }));
      if (!result.ok) throw new Error(result.note);
      // The index still believes in the part: only the sync knows it is gone.
      expect(result.write.replacedPartId).toBeDefined();
      await expect(context.sync()).rejects.toMatchObject({
        code: "ItemNotFound",
      });
      expect(store.adoptDeckWrite(result.write)).toBe(false);
    });
    expect(recordsIn(presentation)).toEqual([]);
  });

  it("writes a record over the record cap as identity only", async () => {
    const huge = buildDeckRecord(
      {
        id: idOf(1),
        kind: "table",
        label: "Model!B4:F12 table",
        token: "token-1",
        sentAt: "2026-10-02T09:00:00.000Z",
      },
      tablePayload(60, 20, () => "x".repeat(400)),
      4,
      "2026-10-03T08:00:00.000Z",
    );
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, huge);
      if (!result.ok) throw new Error(result.note);
      expect(result.write.record).toMatchObject({
        dataKept: false,
        payload: null,
      });
      expect(result.write.bytes).toBeLessThanOrEqual(DECK_RECORD_MAX_BYTES);
      await context.sync();
    });
    expect(recordsIn(presentation)).toMatchObject([
      { id: idOf(1), rev: 4, dataKept: false, payload: null },
    ]);
  });

  it("adds a second part while the index is not to be trusted, and reconcile later removes the old one", async () => {
    await seedParts([xmlOf(1)]);
    await read();
    store.distrustDeckStore();
    await PowerPoint.run(async (context) => {
      const result = store.writeDeckRecord(context, recordOf(1, { rev: 2 }));
      if (!result.ok) throw new Error(result.note);
      expect(result.write.replacedPartId).toBeUndefined();
      await context.sync();
      expect(store.adoptDeckWrite(result.write)).toBe(true);
    });
    expect(recordsIn(presentation).map((one) => one.rev)).toEqual([1, 2]);
    expect(store.deckStoreBytes()).toBe(
      xmlOf(1).length + xmlOf(1, { rev: 2 }).length,
    );
    await read();
    await PowerPoint.run((context) => store.reconcileDuplicates(context));
    expect(recordsIn(presentation).map((one) => one.rev)).toEqual([2]);
  });

  it("refuses a record that is too large to keep even as identity only", async () => {
    const result = await PowerPoint.run((context) =>
      Promise.resolve(
        store.writeDeckRecord(context, {
          ...recordOf(1),
          label: "x".repeat(DECK_RECORD_MAX_BYTES),
        }),
      ),
    );
    expect(!result.ok && result.note).toMatch(/too large to keep$/);
    expect(recordsIn(presentation)).toEqual([]);
  });

  it("says so, and writes nothing, when the host has no custom XML parts", () => {
    const result = store.writeDeckRecord(
      { presentation: {} } as unknown as PowerPoint.RequestContext,
      recordOf(1),
    );
    expect(result.ok).toBe(false);
  });
});

describe("deleting a record", () => {
  it("queues the delete of its part and forgets it", async () => {
    await seedParts([xmlOf(1), xmlOf(2)]);
    await read();
    await PowerPoint.run(async (context) => {
      expect(store.deleteDeckRecord(context, idOf(1))).toBe(true);
      await context.sync();
    });
    expect(recordsIn(presentation)).toEqual([recordOf(2)]);
    expect([...store.heldRecords().keys()]).toEqual([idOf(2)]);
  });

  it("takes the other parts of the same link with it, so no copy of its token stays", async () => {
    await seedParts([xmlOf(1, { rev: 1 }), xmlOf(1, { rev: 2 }), xmlOf(2)]);
    await read();
    await PowerPoint.run(async (context) => {
      expect(store.deleteDeckRecord(context, idOf(1))).toBe(true);
      await context.sync();
    });
    expect(recordsIn(presentation)).toEqual([recordOf(2)]);
    expect(store.deckStoreBytes()).toBe(xmlOf(2).length);
  });

  it("queues the delete of the part it held, which the host refuses at the sync when that part is gone", async () => {
    await seedParts([xmlOf(1)]);
    await read();
    await removeAllParts();
    await PowerPoint.run(async (context) => {
      expect(store.deleteDeckRecord(context, idOf(1))).toBe(true);
      await expect(context.sync()).rejects.toMatchObject({
        code: "ItemNotFound",
      });
    });
    // The caller of a refused batch distrusts the index and reads the deck.
    store.distrustDeckStore();
    await read();
    expect(store.deckStoreStale()).toBe(false);
    expect(store.heldRecords().size).toBe(0);
  });

  it("queues nothing for a link the deck holds no record of", async () => {
    await seedParts([xmlOf(2)]);
    await read();
    const from = helpers.syncCount();
    await PowerPoint.run((context) => {
      expect(store.deleteDeckRecord(context, idOf(1))).toBe(false);
      return Promise.resolve();
    });
    expect(helpers.syncCount()).toBe(from);
    expect(recordsIn(presentation)).toEqual([recordOf(2)]);
  });
});

describe("purging the store", () => {
  it("deletes every part of the namespace, the unreadable ones too, and nothing else", async () => {
    await PowerPoint.run(async (context) => {
      context.presentation.customXmlParts.add(`<other xmlns="urn:other"/>`);
      await context.sync();
    });
    await seedParts([xmlOf(1), xmlOf(2), FOREIGN, NEWER]);
    await read();
    const purged = await PowerPoint.run((context) =>
      store.purgeDeckStore(context),
    );
    expect(purged).toBe(4);
    expect(partsIn(presentation)).toEqual([`<other xmlns="urn:other"/>`]);
    expect(store.heldRecords().size).toBe(0);
    expect(store.deckStoreBytes()).toBe(0);
  });

  it("is one sync on a deck with nothing of ours in it", async () => {
    const from = helpers.syncCount();
    expect(
      await PowerPoint.run((context) => store.purgeDeckStore(context)),
    ).toBe(0);
    expect(helpers.syncCount() - from).toBe(1);
  });
});

describe("whether the deck can keep records at all", () => {
  it("needs PowerPointApi 1.7", () => {
    expect(store.deckStoreAvailable()).toBe(true);
    helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.7"),
    );
    expect(store.deckStoreAvailable()).toBe(false);
  });

  it("is false with no PowerPoint at all, which is what a unit test of another module sees", () => {
    uninstallFakePpt();
    expect(store.deckStoreAvailable()).toBe(false);
  });
});
