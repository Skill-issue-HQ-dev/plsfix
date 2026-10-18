// What Break link and Change source take out of the deck's own store: the
// record goes with the last object that carried the link, never with the
// first, and a change of source moves it to the new link. Strict load
// semantics on.

import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  PNG,
  partsIn,
  recordsIn,
  removeAllParts,
  seedParts,
  withoutCustomXml,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  seedLink,
  seedText,
} from "./ppt.support";
import type * as HostModule from "../src/ppt/host";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let host: typeof HostModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  host = await import("../src/ppt/host");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(presentation, true);

describe("Break link", () => {
  it("deletes the record with the last object that carried the link", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    expect(committed()).toHaveLength(1);
    await host.breakLink(row!.found);
    expect(partsIn(presentation)).toEqual([]);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect([shape.tags.get(TAG_LINK), shape.tags.get(TAG_KEY)]).toEqual([
      undefined,
      undefined,
    ]);
  });

  it("keeps it while another copy of the link is still in the deck", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    expect(committed()).toHaveLength(1);
    await host.breakLink(rows[0]!.found);
    expect(committed()).toHaveLength(1);
    await host.breakLink(rows[1]!.found);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("takes the record out even when a co-author replaced its part, and leaves no tag behind", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    const [xml] = partsIn(presentation);
    // The co-author's pane rewrote the record under a part id this pane has
    // never seen: the id the index holds is gone, the token is still in the deck.
    await removeAllParts();
    await seedParts([xml!]);
    await host.breakLink(row!.found);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect([shape.tags.get(TAG_LINK), shape.tags.get(TAG_KEY)]).toEqual([
      undefined,
      undefined,
    ]);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("keeps it while another copy is in the deck that this pane did not put there", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    // The pane is opened again on the same deck: it has seen no insert.
    vi.resetModules();
    links = await import("../src/ppt/links");
    host = await import("../src/ppt/host");
    await (await import("../src/ppt/deck-boot")).loadDeckStore();
    const rows = await links.listLinks(relay);
    await host.breakLink(rows[0]!.found);
    expect(committed()).toHaveLength(1);
    await host.breakLink(rows[1]!.found);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("leaves another link's record alone", async () => {
    const first = await seedText("EUR 3.2m");
    const second = await seedLink(PNG);
    await links.insertFromInbox(first, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(second, ws, relay);
    const rows = await links.listLinks(relay);
    expect(committed()).toHaveLength(2);
    await host.breakLink(rows[0]!.found);
    expect(committed().map((one) => one.id)).toEqual([second.id]);
  });

  it("takes the record out in the batch that drops the tags, a round trip no dearer than a plain break", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    const from = helpers.syncCount();
    await host.breakLink(row!.found);
    expect(
      helpers.syncCount() - from,
      "the record's delete rides the sync that drops the tags",
    ).toBe(1);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("spends one round trip on every copy it lets go, the record leaving with the last", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    const rows = await links.listLinks(relay);
    let from = helpers.syncCount();
    await host.breakLink(rows[0]!.found);
    expect(helpers.syncCount() - from, "the first copy's tags only").toBe(1);
    expect(committed()).toHaveLength(1);
    from = helpers.syncCount();
    await host.breakLink(rows[1]!.found);
    expect(
      helpers.syncCount() - from,
      "the last copy's tags and the record, together",
    ).toBe(1);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("writes the next update into the one part it still has after a break the host refused twice", async () => {
    const item = await seedLink(PNG);
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    const tags = [shape.tags.get(TAG_LINK)!, shape.tags.get(TAG_KEY)!];
    helpers.failNextSync();
    helpers.failNextSync(undefined, 1);
    await expect(host.breakLink(row!.found)).rejects.toThrow();
    // A real host that refuses the batch applies none of it: the tags stay.
    shape.tags.set(TAG_LINK, tags[0]!);
    shape.tags.set(TAG_KEY, tags[1]!);
    await pushAgain(item, fakePng(1600, 800));
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(partsIn(presentation)).toHaveLength(1);
    expect(committed()[0]).toMatchObject({ id: item.id, rev: 2 });
  });

  it("reads the deck before it deletes anything while the index is in doubt", async () => {
    const store = await import("../src/ppt/deck-store");
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    store.distrustDeckStore();
    const from = helpers.syncCount();
    await host.breakLink(row!.found);
    expect(
      helpers.syncCount() - from,
      "the tags; the deck read (two); the delete, none of it in the tags' batch",
    ).toBe(4);
    expect(partsIn(presentation)).toEqual([]);
  });

  it("takes a link's record out again when the link comes back and is broken a second time", async () => {
    const item = await seedText("EUR 3.2m");
    const from = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    const spent = helpers.syncCount() - from;
    const [row] = await links.listLinks(relay);
    await host.breakLink(row!.found);
    expect(partsIn(presentation)).toEqual([]);

    helpers.selectSlide(presentation.slides[1]!.id);
    const again = helpers.syncCount();
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(
      helpers.syncCount() - again,
      "a second insert of the link costs what the first did",
    ).toBe(spent);
    expect(committed().map((one) => one.id)).toEqual([item.id]);
    // No scan in between: the pane has only its own count of what the deck holds.
    await host.breakLink({
      ...row!.found,
      slideId: placed.slideId,
      shapeId: placed.shapeId,
    });
    expect(partsIn(presentation)).toEqual([]);
  });

  it("leaves no record behind when the host applies the batch and then refuses it", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    const [row] = await links.listLinks(relay);
    helpers.refuseNextSync();
    await host.breakLink(row!.found);
    const shape = presentation.slides[0]!.shapes[0]!;
    expect(shape.tags.get(TAG_LINK)).toBeUndefined();
    expect(partsIn(presentation)).toEqual([]);
  });

  it("spends no round trip on a deck that holds no record of the link", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.setSupported(() => true);
    const [row] = await links.listLinks(relay);
    expect(partsIn(presentation)).toEqual([]);
    const from = helpers.syncCount();
    await host.breakLink(row!.found);
    expect(helpers.syncCount() - from, "the tag delete is the only sync").toBe(
      1,
    );
  });
});

describe("Change source", () => {
  it("writes the new link's record and deletes the old one's", async () => {
    const { changeSource } = await import("../src/ppt/change-source");
    const old = await seedLink(PNG);
    await links.insertFromInbox(old, ws, relay);
    const next = await seedLink(fakePng(1600, 800), "Model_v5.xlsx");
    const [row] = await links.listLinks(relay);
    await changeSource(row!, next, ws, relay);
    expect(committed().map((one) => one.id)).toEqual([next.id]);
    expect(committed()[0]).toMatchObject({
      rev: 1,
      token: next.token,
      src: { workbook: "Model_v5.xlsx" },
    });
  });

  it("keeps the old link's record when the repaint fails and the old tags are put back", async () => {
    const { changeSource } = await import("../src/ppt/change-source");
    const old = await seedLink(PNG);
    await links.insertFromInbox(old, ws, relay);
    const next = await seedLink(fakePng(1600, 800), "Model_v5.xlsx");
    const [row] = await links.listLinks(relay);
    // The retag lands; the repaint's batch does not, nor does anything the
    // pane tries after it, until the old tags are put back.
    for (const after of [1, 2, 3, 4]) helpers.failNextSync(undefined, after);
    await expect(changeSource(row!, next, ws, relay)).rejects.toThrow();
    expect(committed().map((one) => one.id)).toEqual([old.id]);
  });
});
