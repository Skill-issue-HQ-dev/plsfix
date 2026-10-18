// What the link flows leave in the deck's own store: an insert writes the
// link's record, Update all and Revert rewrite it, and none of it costs a flow
// a round trip. A deck with only tags gains its records on the first Update
// all. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  customXmlCommitted,
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  KINDS,
  NEXT_CELLS,
  PNG,
  partsIn,
  recordsIn,
  withoutCustomXml,
  type Kind,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushAgain,
  pushText,
  seedChart,
  seedLink,
  seedText,
  src,
} from "./ppt.support";
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

describe("an insert", () => {
  it.each(KINDS)(
    "leaves the record of $name in the deck",
    async ({ seed, kind, kept }) => {
      const item = await seed();
      expect(customXmlCommitted(presentation)).toEqual([]);
      await links.insertFromInbox(item, ws, relay);
      expect(committed()).toHaveLength(1);
      expect(committed()[0]).toMatchObject({
        v: 1,
        id: item.id,
        kind,
        label: item.label,
        src,
        token: item.token,
        rev: 1,
        sentAt: item.createdAt,
        dataKept: kept,
      });
    },
  );

  it("keeps the data a deck needs to draw the object again, and not its picture", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    const [record] = committed();
    expect(record?.payload).toMatchObject({
      kind: "picture",
      width: 800,
      height: 400,
      chart: columnChart(6),
    });
    expect(JSON.stringify(partsIn(presentation))).not.toContain(PNG);
  });

  it("keeps the project the link was sent under", async () => {
    const item = { ...(await seedText("EUR 3.2m")), project: "Amasty" };
    await links.insertFromInbox(item, ws, relay);
    expect(committed()[0]).toMatchObject({ project: "Amasty" });
  });

  it("keeps one part for a link inserted twice, one copy on each of two slides", async () => {
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(item, ws, relay);
    expect(partsIn(presentation)).toHaveLength(1);
  });

  it.each(KINDS)(
    "costs no more round trips for $name than a deck that keeps no records",
    async ({ seed }) => {
      const spent = async (keeps: boolean): Promise<number> => {
        ({ links, helpers, relay } = await bootPpt());
        if (!keeps) withoutCustomXml(helpers);
        const item = await seed();
        const from = helpers.syncCount();
        await links.insertFromInbox(item, ws, relay);
        return helpers.syncCount() - from;
      };
      const without = await spent(false);
      const keeping = await spent(true);
      expect(
        keeping,
        "the record rides the batch that tags the object: no round trip of its own",
      ).toBe(without);
    },
  );
});

describe("Update all", () => {
  async function updated(kind: Kind): Promise<InboxItem> {
    const item = await kind.seed();
    await links.insertFromInbox(item, ws, relay);
    await kind.push(item);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    return item;
  }

  it.each(KINDS)(
    "bumps the revision of $name in the part it already has",
    async (kind) => {
      const item = await updated(kind);
      expect(partsIn(presentation)).toHaveLength(1);
      expect(committed()[0]).toMatchObject({
        id: item.id,
        rev: 2,
        label: item.label,
        sentAt: item.createdAt,
        dataKept: kind.kept,
      });
    },
  );

  it("keeps the label and the first-sent time the link was inserted with", async () => {
    const sent = {
      ...(await seedText("EUR 3.2m")),
      label: "Revenue bridge",
      createdAt: "2026-10-01T08:00:00.000Z",
    };
    await links.insertFromInbox(sent, ws, relay);
    await pushText(sent, "EUR 3.9m");
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(committed()[0]).toMatchObject({
      rev: 2,
      label: "Revenue bridge",
      sentAt: "2026-10-01T08:00:00.000Z",
    });
  });

  it("brings the new figures into the record of a table", async () => {
    await updated(KINDS[1]!);
    expect(committed()[0]?.payload).toMatchObject({
      kind: "table",
      cells: NEXT_CELLS,
    });
  });

  it("stamps the record with the time of the push it holds", async () => {
    const item = await updated(KINDS[0]!);
    const tag = presentation.slides[0]!.shapes[0]!.tags.get(TAG_LINK)!;
    const pushedAt = (JSON.parse(tag) as { pushedAt: string }).pushedAt;
    expect(committed()[0]).toMatchObject({ id: item.id, pushedAt });
  });

  it.each(KINDS)(
    "costs no more round trips for $name than a deck that keeps no records",
    async (kind) => {
      const spent = async (keeps: boolean): Promise<number> => {
        ({ links, helpers, relay } = await bootPpt());
        if (!keeps) withoutCustomXml(helpers);
        const item = await kind.seed();
        await links.insertFromInbox(item, ws, relay);
        await kind.push(item);
        const rows = await links.listLinks(relay);
        const from = helpers.syncCount();
        await links.updateLinks(rows, relay);
        return helpers.syncCount() - from;
      };
      const without = await spent(false);
      const keeping = await spent(true);
      expect(
        keeping,
        "the record rides the repaint batch: no round trip of its own",
      ).toBe(without);
    },
  );

  it("gives a deck that holds only tags its records on the first Update all", async () => {
    withoutCustomXml(helpers);
    const text = await seedText("EUR 3.2m");
    const range = await seedLink(PNG);
    await links.insertFromInbox(text, ws, relay);
    helpers.selectSlide(presentation.slides[1]!.id);
    await links.insertFromInbox(range, ws, relay);
    expect(partsIn(presentation)).toEqual([]);

    helpers.setSupported(() => true);
    await pushText(text, "EUR 3.9m");
    await pushAgain(range, fakePng(1600, 800));
    await links.updateLinks(await links.listLinks(relay), relay);

    const records = committed();
    expect(records.map((one) => [one.id, one.rev, one.dataKept])).toEqual([
      [text.id, 2, true],
      [range.id, 2, false],
    ]);
    // Nothing remembered a label for them, so the source says what it is.
    expect(records[0]?.label).toBe("Model!B4:F12 text");
    expect(records[0]?.sentAt).toBe(records[0]?.pushedAt);
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(partsIn(presentation)).toHaveLength(2);
  });

  it("writes the older revision back into the record when Revert goes one step back", async () => {
    const { revertLinks } = await import("../src/ppt/revert");
    const item = await seedText("EUR 3.2m");
    await links.insertFromInbox(item, ws, relay);
    await pushText(item, "EUR 3.9m");
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(committed()[0]).toMatchObject({ rev: 2 });
    const rows = await links.listLinks(relay);
    await revertLinks(rows, relay);
    expect(committed()[0]).toMatchObject({
      rev: 1,
      payload: { kind: "text", text: "EUR 3.2m" },
    });
    expect(partsIn(presentation)).toHaveLength(1);
  });
});
