// A deck from before the deck store holds tags and no records. Its first Update
// all finds every link current and repaints nothing, yet the deck gains its
// records: one batch, one sync, no repaint. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { InboxItem } from "../src/link/model";
import { RelayError } from "../src/link/relay-error";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  fillTheDeck,
  KINDS,
  partsIn,
  recordsIn,
  withoutCustomXml,
} from "./ppt.deck-store.support";
import { bootPpt, memoryStore, pushText, seedText } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(presentation, true);

// Everything a repaint would have touched: each shape's id and how often its
// picture was set.
const painted = (): [string, number][] =>
  presentation.slides.flatMap((slide) =>
    slide.shapes.map((shape): [string, number] => [
      shape.id,
      shape.setImageCalls,
    ]),
  );

// One link of every kind in a deck the host could not keep records for, which
// can keep them again by the time Update all runs.
async function tagsOnlyDeck(): Promise<InboxItem[]> {
  withoutCustomXml(helpers);
  const items: InboxItem[] = [];
  for (const [at, kind] of KINDS.entries()) {
    helpers.selectSlide(presentation.slides[at % 3]!.id);
    const item = await kind.seed();
    await links.insertFromInbox(item, ws, relay);
    items.push(item);
  }
  helpers.setSupported(() => true);
  expect(partsIn(presentation)).toEqual([]);
  return items;
}

describe("a deck that holds only tags, every link current", () => {
  it("gains all its records on the first Update all, with no repaint and one save sync", async () => {
    const items = await tagsOnlyDeck();
    const rows = await links.listLinks(relay);
    const before = painted();
    const from = helpers.syncCount();

    const summary = await links.updateLinks(rows, relay);

    expect(summary).toMatchObject({ updated: 0, current: 5, failed: 0 });
    expect(summary.notes).toEqual([]);
    expect(helpers.syncCount() - from, "one batch saves every record").toBe(1);
    expect(painted()).toEqual(before);
    expect(new Set(committed().map((one) => one.id))).toEqual(
      new Set(items.map((one) => one.id)),
    );
    expect(committed().map((one) => one.rev)).toEqual([1, 1, 1, 1, 1]);
    expect(
      committed()
        .map((one) => one.dataKept)
        .sort(),
    ).toEqual([false, false, true, true, true]);
  });

  it("gains the record through the per-row route too, when the relay refuses the batch", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.setSupported(() => true);
    Object.assign(relay, {
      fetchLinks: () =>
        Promise.reject(
          new RelayError("server", "relay POST /api/links/fetch: 404", 404),
        ),
    });
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(committed()).toMatchObject([{ id: item.id, rev: 1 }]);
  });

  it("still repaints a link whose source moved on, and gives it its record in the same batch", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.setSupported(() => true);
    await pushText(item, "EUR 3.9m");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, current: 0, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.text).toBe("EUR 3.9m");
    expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
  });

  it("counts both copies of a link that has its record as updated, one of them already at the new revision", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 3.9m");
    const rows = await links.listLinks(relay);
    await links.updateLinks([rows[0]!], relay);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 2, current: 0, failed: 0 });
    expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
  });

  it("asks the relay nothing more and writes nothing on the next Update all", async () => {
    await tagsOnlyDeck();
    await links.updateLinks(await links.listLinks(relay), relay);
    const rows = await links.listLinks(relay);
    const from = helpers.syncCount();
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 0, current: 5, failed: 0 });
    expect(helpers.syncCount()).toBe(from);
    expect(partsIn(presentation)).toHaveLength(5);
  });

  it("saves one record for a link that sits on two slides", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    helpers.setSupported(() => true);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 0, current: 2, failed: 0 });
    expect(committed()).toMatchObject([{ id: item.id, rev: 1 }]);
  });

  it("says in the summary what a full deck could not keep, and the rows stay current", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.setSupported(() => true);
    await fillTheDeck();
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(summary.notes).toEqual([
      "Model!B4:F12 text The deck's copy of this link was not saved: the link store in this deck is full",
    ]);
    expect(committed()).toEqual([]);
  });

  it("leaves a deck on a host without custom XML parts exactly as it was", async () => {
    withoutCustomXml(helpers);
    await seedText("EUR 3.2m").then((item) =>
      links.insertFromInbox(item, ws, relay),
    );
    const rows = await links.listLinks(relay);
    const from = helpers.syncCount();
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 0, current: 1, failed: 0 });
    expect(helpers.syncCount()).toBe(from);
  });
});
