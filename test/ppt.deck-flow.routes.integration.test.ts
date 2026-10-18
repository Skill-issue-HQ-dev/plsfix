// The deck store on the routes a link does not usually take: a host with
// custom XML parts but no PowerPointApi 1.8 (the picture goes in through the
// selection), and what the deck cannot keep at all. The chart routes are in
// ppt.deck-flow.charts.integration.test.ts. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  CELLS,
  fillTheDeck,
  PNG,
  partsIn,
  recordsIn,
  withoutCustomXml,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  pushTable,
  seedLink,
  seedTable,
  seedText,
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
  vi.useRealTimers();
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(presentation, true);

type Profile = (set: string, version: string) => boolean;
// Custom XML parts (1.7) and nothing newer: no fill.setImage, no groups.
const upTo17: Profile = (set, version) =>
  set !== "PowerPointApi" || Number(version.split(".")[1]) <= 7;

// What `act` costs in round trips on a fresh deck of this sort of host, with
// custom XML parts and without; `act` counts its own.
async function costs(
  profile: Profile,
  act: () => Promise<number>,
): Promise<{ keeping: number; without: number }> {
  const spent = async (keeps: boolean): Promise<number> => {
    ({ links, helpers, relay, presentation } = await bootPpt());
    helpers.setSupported(
      (set, version) =>
        profile(set, version) &&
        (keeps || !(set === "PowerPointApi" && version === "1.7")),
    );
    return act();
  };
  return { keeping: await spent(true), without: await spent(false) };
}

describe("a host with custom XML parts and no PowerPointApi 1.8", () => {
  it("keeps the record of a picture that went in through the selection, and of its repaint", async () => {
    helpers.setSupported(upTo17);
    const item = await seedLink(PNG);
    await links.insertFromInbox(item, ws, relay);
    expect(committed()).toMatchObject([{ id: item.id, rev: 1 }]);
    await pushAgain(item, fakePng(1600, 800));
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(committed()).toMatchObject([{ id: item.id, rev: 2 }]);
    expect(partsIn(presentation)).toHaveLength(1);
  });

  it("costs the insert and the repaint no more round trips than a host without them", async () => {
    const act = async (): Promise<number> => {
      const item = await seedLink(PNG);
      const from = helpers.syncCount();
      await links.insertFromInbox(item, ws, relay);
      await pushAgain(item, fakePng(1600, 800));
      await links.updateLinks(await links.listLinks(relay), relay);
      return helpers.syncCount() - from;
    };
    const { keeping, without } = await costs(upTo17, act);
    expect(keeping, "the record rides the batch that tags the picture").toBe(
      without,
    );
  });
});

describe("what the deck cannot keep", () => {
  it("writes a table too big for a record as identity only, and the insert is whole", async () => {
    const cells = Array.from({ length: 60 }, () =>
      Array.from({ length: 20 }, () => ({ t: "x".repeat(400) })),
    );
    const item = await seedTable(
      cells,
      Array.from({ length: 20 }, () => 20),
    );
    await links.insertFromInbox(item, ws, relay);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(committed()).toMatchObject([
      { id: item.id, dataKept: false, payload: null, rev: 1 },
    ]);
  });

  it("refuses a write the deck has no room for, says so in the insert's note, and inserts", async () => {
    await fillTheDeck();
    const item = await seedText("EUR 3.2m");
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(placed.note).toBe(
      "The deck's copy of this link was not saved: the link store in this deck is full",
    );
    expect(links.insertNote(placed)).toContain("was not saved");
    expect(committed()).toEqual([]);
  });

  it("keeps only the identity of a table the deck has no room for whole, says so, and inserts", async () => {
    await fillTheDeck(3000);
    const cells = Array.from({ length: 20 }, () =>
      Array.from({ length: 10 }, () => ({ t: "12 345" })),
    );
    const item = await seedTable(
      cells,
      Array.from({ length: 10 }, () => 40),
    );
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(presentation.slides[0]!.shapes[0]!.type).toBe("Table");
    expect(placed.note).toBe(
      "The deck keeps this link's identity but not its data: the link store in this deck is full",
    );
    expect(committed()).toMatchObject([
      { id: item.id, rev: 1, dataKept: false, payload: null },
    ]);
  });

  it("refuses a table with the same note when the deck is full, and the table is whole", async () => {
    await fillTheDeck();
    const item = await seedTable(CELLS, [90, 60]);
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toMatch(/^The deck's copy of this link was not saved/);
    expect(presentation.slides[0]!.shapes[0]!.type).toBe("Table");
  });

  // A deck that holds only tags, on a host that can keep records, whose store
  // is full: the repaint lands and says, row by row, what it could not keep.
  it.each([
    ["a picture, repainted in the one batch", () => seedLink(PNG)],
    ["a table, repainted on its own", () => seedTable(CELLS, [90, 60])],
  ])(
    "tells Update all which copy it could not keep: %s",
    async (_name, seed) => {
      withoutCustomXml(helpers);
      const item = await seed();
      await links.insertFromInbox(item, ws, relay);
      helpers.setSupported(() => true);
      await fillTheDeck();
      await (item.kind === "table"
        ? pushTable(item, CELLS, [90, 60])
        : pushAgain(item, fakePng(1600, 800)));

      const summary = await links.updateLinks(
        await links.listLinks(relay),
        relay,
      );

      expect(summary).toMatchObject({ updated: 1, failed: 0 });
      expect(summary.notes).toHaveLength(1);
      expect(summary.notes[0]).toMatch(
        /The deck's copy of this link was not saved: the link store in this deck is full$/,
      );
      expect(committed()).toEqual([]);
    },
  );

  it("inserts and updates as it always did on a host without custom XML parts", async () => {
    withoutCustomXml(helpers);
    const item = await seedText("EUR 3.2m");
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toBeUndefined();
    expect(partsIn(presentation)).toEqual([]);
    const deckBoot = await import("../src/ppt/deck-boot");
    expect(deckBoot.deckRecords().size).toBe(0);
  });

  it("inserts, with a note, on a host that claims custom XML parts and has none", async () => {
    uninstallFakePpt();
    ({ presentation, helpers } = installFakePpt({
      slides: 3,
      strictLoad: false,
    }));
    helpers.selectSlide(presentation.slides[0]!.id);
    const run = PowerPoint.run.bind(PowerPoint);
    const broken = (
      callback: (context: PowerPoint.RequestContext) => Promise<unknown>,
    ): Promise<unknown> =>
      run((context: PowerPoint.RequestContext) => {
        Object.defineProperty(context.presentation, "customXmlParts", {
          value: undefined,
        });
        return callback(context);
      });
    (PowerPoint as unknown as { run: typeof broken }).run = broken;
    const item = await seedText("EUR 3.2m");
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(placed.note).toMatch(
      /^The deck's copy of this link was not saved: /,
    );
  });
});
