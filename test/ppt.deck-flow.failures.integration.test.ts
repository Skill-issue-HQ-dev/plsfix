// When a host does not take a record: a batch it refuses with a deck write in
// it, one it swallows, a part gone from under the pane, a save it refuses
// twice. A record never fails a flow the host would have run without it, and a
// flow that cannot keep one says so in its note. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import { settleHungSync } from "./hung-sync";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  CELLS,
  KINDS,
  NEXT_CELLS,
  partsIn,
  recordsIn,
  removeAllParts,
  WIDTHS,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  pushTable,
  pushText,
  seedLink,
  seedTable,
  seedText,
} from "./ppt.support";
import type * as FlowModule from "../src/ppt/deck-flow";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);

let links: typeof LinksModule;
let flow: typeof FlowModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  flow = await import("../src/ppt/deck-flow");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(presentation, true);
const refused = new Error("GeneralException: the host refused");

// How many syncs one text insert spends on this fake, deck store and all: the
// last of them is the batch that adds the box, tags it and carries its record.
async function textInsertSyncs(): Promise<number> {
  const probe = await seedText("probe");
  const from = helpers.syncCount();
  await links.insertFromInbox(probe, ws, relay);
  return helpers.syncCount() - from;
}

describe("an insert whose batch the host refuses", () => {
  it("fails as it always did, leaves no box and no record, and the deck stops riding", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    flow = await import("../src/ppt/deck-flow");
    const item = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);

    await expect(links.insertFromInbox(item, ws, relay)).rejects.toThrow(
      "the host refused",
    );

    expect(presentation.slides[0]!.shapes).toHaveLength(0);
    expect(partsIn(presentation)).toEqual([]);
    expect(flow.deckRiding()).toBe(false);
  });

  it("goes in on the next try, its record saved in a batch of its own after the deck is read again", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    const first = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);
    await expect(links.insertFromInbox(first, ws, relay)).rejects.toThrow();

    const again = await seedText("EUR 3.9m");
    const from = helpers.syncCount();
    const placed = await links.insertFromInbox(again, ws, relay);

    expect(placed.note).toBeUndefined();
    expect(committed().map((one) => one.id)).toEqual([again.id]);
    expect(
      helpers.syncCount() - from,
      "the flow's own syncs, one to read the deck again, one to save the record",
    ).toBe(syncs + 2);
  });

  it("does not stop the deck riding when the box a row names is simply gone", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    presentation.deleteShape(rows[0]!.found.shapeId);
    await pushText(item, "EUR 3.9m");
    const summary = await links.updateLinks(rows, relay);
    expect(summary.failed).toBe(1);
    expect(flow.deckRiding()).toBe(true);
  });
});

describe("an Update all batch the host refuses with records in it", () => {
  async function twoPictures(): Promise<LinksModule.LinkRow[]> {
    const first = await seedLink(PNG);
    await links.insertFromInbox(first, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    const second = await seedLink(PNG);
    await links.insertFromInbox(second, ws, relay);
    await pushAgain(first, fakePng(1600, 800));
    await pushAgain(second, fakePng(1600, 800));
    return links.listLinks(relay);
  }

  it("is run again row by row as it always was, without the records, which are saved on their own", async () => {
    const rows = await twoPictures();
    helpers.failNextSync(refused);
    const from = helpers.syncCount();

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect(committed().map((one) => one.rev)).toEqual([2, 2]);
    expect(partsIn(presentation)).toHaveLength(2);
    expect(flow.deckRiding()).toBe(false);
    expect(
      helpers.syncCount() - from,
      "the refused batch; the first row; the deck read again (two); its record; the second row; its record",
    ).toBe(7);
  });

  it("keeps riding when it was the shape and not the record that the host refused", async () => {
    const rows = await twoPictures();
    presentation.deleteShape(rows[0]!.found.shapeId);
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, failed: 1 });
    expect(flow.deckRiding()).toBe(true);
    expect(committed().map((one) => one.rev)).toEqual([1, 2]);
  });
});

describe("a Break whose batch carries the record's delete", () => {
  async function inserted(): Promise<LinksModule.LinkRow> {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    return (await links.listLinks(relay))[0]!;
  }

  it("is run once more without the record when the host refuses it, and goes through", async () => {
    const host = await import("../src/ppt/host");
    const row = await inserted();
    const from = helpers.syncCount();
    helpers.failNextSync(refused);
    await host.breakLink(row.found);
    expect(
      helpers.syncCount() - from,
      "the refused batch; the plain break; the deck read again (two); the record's delete",
    ).toBe(5);
    expect(partsIn(presentation)).toEqual([]);
    expect(flow.deckRiding()).toBe(true);
  });

  it("fails with the host's own words when the second try is refused too, and leaves the record", async () => {
    const host = await import("../src/ppt/host");
    const row = await inserted();
    helpers.failNextSync(refused);
    helpers.failNextSync(refused, 1);
    await expect(host.breakLink(row.found)).rejects.toThrow("the host refused");
    expect(committed()).toHaveLength(1);
  });

  it("takes the record out in a batch of its own once a refused insert has stopped the deck riding", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    flow = await import("../src/ppt/deck-flow");
    const host = await import("../src/ppt/host");
    const first = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);
    await expect(links.insertFromInbox(first, ws, relay)).rejects.toThrow();
    expect(flow.deckRiding()).toBe(false);
    // The deck is read again and the link's record is saved on its own.
    const item = await seedText("EUR 3.9m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    const from = helpers.syncCount();
    await host.breakLink(row!.found);
    expect(
      helpers.syncCount() - from,
      "the tags, then the record in a batch of its own",
    ).toBe(2);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("does not stop the deck riding when the host never answers", async () => {
    const host = await import("../src/ppt/host");
    const row = await inserted();
    helpers.hangNextSync();
    vi.useFakeTimers();
    await expect(settleHungSync(host.breakLink(row.found))).rejects.toThrow();
    expect(flow.deckRiding()).toBe(true);
  });

  it("does not stop the deck riding when the box a row names is simply gone", async () => {
    const host = await import("../src/ppt/host");
    const row = await inserted();
    presentation.deleteShape(row.found.shapeId);
    await expect(host.breakLink(row.found)).rejects.toThrow();
    expect(flow.deckRiding()).toBe(true);
    expect(committed()).toHaveLength(1);
  });
});

describe("a record that rode an Update all the host refuses", () => {
  // A co-author removed the part the index still holds, so the rewrite that
  // rides the repaint's batch is the op the host cannot do.
  for (const kind of KINDS) {
    it(`is not the row's failure for ${kind.name}: it goes through and the record is written again`, async () => {
      const item = await kind.seed();
      await links.insertFromInbox(item, ws, relay);
      await removeAllParts();
      await kind.push(item);
      const summary = await links.updateLinks(
        await links.listLinks(relay),
        relay,
      );
      expect(summary).toMatchObject({ updated: 1, failed: 0 });
      expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
      expect(flow.deckRiding()).toBe(true);
    });
  }

  it("is not the row's failure for a table whose grid grew: it is built again and the record is written", async () => {
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay);
    await removeAllParts();
    await pushTable(
      item,
      [...NEXT_CELLS, [{ t: "Profit" }, { t: "540" }]],
      WIDTHS,
    );
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
  });

  it("is not run again when the host never answers, and the deck keeps riding", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 3.9m");
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();
    helpers.hangNextSync();
    vi.useFakeTimers();
    const summary = await settleHungSync(links.updateLinks(rows, relay));
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(helpers.syncCount()).toBe(before + 1);
    expect(flow.deckRiding()).toBe(true);
  });

  it("leaves nothing staged when the second try fails too: the next update writes the record", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    await removeAllParts();
    await pushText(item, "EUR 3.9m");
    const rows = await links.listLinks(relay);
    helpers.failNextSync(refused);
    helpers.failNextSync(refused, 1);
    await links.updateLinks(rows, relay);
    await pushText(item, "EUR 4.0m");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(committed()).toMatchObject([{ id: item.id, rev: 3 }]);
  });

  it("says nothing about the record in the row's result when the second try goes through", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    await removeAllParts();
    await pushText(item, "EUR 3.9m");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary.notes).toEqual([]);
    expect(summary.failures).toEqual([]);
    expect(presentation.slides[0]!.shapes[0]!.text).toBe("EUR 3.9m");
  });

  it("is run again only once: a second refusal fails the row with the host's own words", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    await removeAllParts();
    await pushText(item, "EUR 3.9m");
    const rows = await links.listLinks(relay);
    helpers.failNextSync(refused);
    helpers.failNextSync(refused, 1);
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
  });

  it("is not run again when no record rode it: the refusal was not the record's", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    flow = await import("../src/ppt/deck-flow");
    const first = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);
    await expect(links.insertFromInbox(first, ws, relay)).rejects.toThrow();
    expect(flow.deckRiding(), "a refused insert stopped the riding").toBe(
      false,
    );

    const item = await seedText("EUR 3.9m");
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 4.0m");
    const rows = await links.listLinks(relay);
    helpers.failNextSync(refused);
    const summary = await links.updateLinks(rows, relay);

    expect(
      summary,
      "one refusal and one attempt: nothing to leave out",
    ).toMatchObject({
      updated: 0,
      failed: 1,
    });
    expect(flow.deckRiding()).toBe(false);
  });
});

describe("an Update all batch the host never answers", () => {
  it("asks for nothing more, keeps the old records, and the next update saves them again", async () => {
    const first = await seedLink(PNG);
    await links.insertFromInbox(first, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    const second = await seedLink(PNG);
    await links.insertFromInbox(second, ws, relay);
    await pushAgain(first, fakePng(1600, 800));
    await pushAgain(second, fakePng(1600, 800));
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();

    helpers.hangNextSync();
    vi.useFakeTimers();
    const summary = await settleHungSync(links.updateLinks(rows, relay));

    expect(summary).toMatchObject({ updated: 0, failed: 2 });
    expect(helpers.syncCount()).toBe(before + 1);
    expect(committed().map((one) => one.rev)).toEqual([1, 1]);
    expect(flow.deckRiding()).toBe(true);

    vi.useRealTimers();
    await pushAgain(first, fakePng(10, 10));
    await pushAgain(second, fakePng(10, 10));
    const next = await links.updateLinks(await links.listLinks(relay), relay);
    expect(next).toMatchObject({ updated: 2, failed: 0 });
    expect(committed().map((one) => one.rev)).toEqual([3, 3]);
    expect(partsIn(presentation)).toHaveLength(2);
  });
});

describe("a save that meets a part gone from under the pane", () => {
  it("is asked again once the deck is read again, and goes in", async () => {
    const store = await import("../src/ppt/deck-store");
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const held = store.heldRecordOf(item.id)!;
    await removeAllParts();
    const from = helpers.syncCount();
    const notes = await flow.saveRecords([{ ...held, rev: 2 }]);
    expect([...notes]).toEqual([]);
    expect(
      helpers.syncCount() - from,
      "the refused save; the deck read again (one: it holds no part); the save",
    ).toBe(3);
    expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
  });
});

describe("a save the host never answers", () => {
  it("is not asked again, and the insert says the copy was not saved", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    const first = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);
    await expect(links.insertFromInbox(first, ws, relay)).rejects.toThrow();

    const item = await seedText("EUR 3.9m");
    const from = helpers.syncCount();
    // The flow's own syncs, the deck read again, then the save never returns.
    helpers.hangNextSync(syncs + 1);
    vi.useFakeTimers();
    const placed = await settleHungSync(links.insertFromInbox(item, ws, relay));
    vi.useRealTimers();

    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(placed.note).toMatch(
      /^The deck's copy of this link was not saved: PowerPoint stopped answering/,
    );
    expect(helpers.syncCount() - from).toBe(syncs + 2);
    expect(committed()).toEqual([]);
  });
});

describe("a save the host refuses twice", () => {
  it("says so in the insert's note and leaves the box where it is", async () => {
    const syncs = await textInsertSyncs();
    ({ links, presentation, helpers, relay } = await bootPpt());
    const first = await seedText("EUR 3.2m");
    helpers.failNextSync(refused, syncs - 1);
    await expect(links.insertFromInbox(first, ws, relay)).rejects.toThrow();

    const item = await seedText("EUR 3.9m");
    // The deck is read again and the save follows; the host refuses the save,
    // is read again, and refuses it once more.
    for (let after = syncs + 1; after <= syncs + 8; after += 1) {
      helpers.failNextSync(refused, after);
    }
    const placed = await links.insertFromInbox(item, ws, relay);

    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(placed.note).toBe(
      "The deck's copy of this link was not saved: GeneralException: the host refused",
    );
    expect(committed()).toEqual([]);
  });
});
