// @vitest-environment jsdom
// L3, the status floor: a deck opened on a wiped or another computer still
// answers for its links from the records it carries, instead of "Not pasted
// yet" or "Source missing". The transport answers first, so newer data and
// every error keep their own words; the deck answers only "nothing there",
// only for a record that holds the data, and only at the rev the slide shows.
// Strict load semantics, over the fake deck, the fake relay and a local store.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, seal } from "../src/link/crypto";
import { LOCAL_REV_BASE } from "../src/link/local";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import { encodeDeckRecord } from "../src/link/deck-record";
import { encodePayload, TAG_LINK, type InboxItem } from "../src/link/model";
import { RelayError } from "../src/link/relay";
import { relayReason } from "../src/link/relay-reason";
import { TRANSPORT_STORAGE_KEY } from "../src/link/transport-setting";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { relativeTime } from "../src/ui/time";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  recordsIn,
  removeAllParts,
  seedParts,
  textPayload,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  memoryStore,
  pushText,
  seedLink,
  seedText,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const MOMENT = new Date(2026, 9, 2, 14, 2, 0);
const MOMENT_WORDS = "data from 02.10 14:02";

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(MOMENT);
  ({ links, presentation, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

// What a pane on THIS computer runs its flows against: relay mode over the
// fake relay, or local mode over a store that holds nothing.
async function transportIn(mode: "local" | "relay") {
  const { openPptTransport } = await import("../src/ppt/transport");
  const keyStore = memoryStore();
  await keyStore.set(TRANSPORT_STORAGE_KEY, mode);
  return openPptTransport({
    keyStore,
    remote: relay,
    paired: () => ws,
    persistence: async () => memoryPersistence(),
  });
}

async function inserted(item: InboxItem): Promise<InboxItem> {
  await links.insertFromInbox(item, ws, relay);
  return item;
}

// The shapes of a deck on a stub host: what must NOT happen is a repaint.
async function quietHost() {
  const real = await import("../src/ppt/host");
  return { ...real, refreshLink: vi.fn(), refreshLinks: vi.fn() };
}

// The Linked objects table as the pane would draw these rows.
async function table(rows: Awaited<ReturnType<typeof links.listLinks>>) {
  const { toRowViews } = await import("../src/ppt/actions");
  const { renderLinkRows } = await import("../src/ppt/views");
  const body = document.createElement("tbody");
  renderLinkRows(body, toRowViews(rows, new Set()), () => undefined);
  return [...body.querySelectorAll<HTMLTableRowElement>("tr[data-key]")].map(
    (row) => ({
      badge: row.querySelector(".badge")?.textContent ?? "",
      meta: row.querySelector(".link-status .link-meta")?.textContent ?? "",
      updated: row.querySelector(".link-updated")?.textContent ?? "",
    }),
  );
}

function recordRev(): number | undefined {
  return recordsIn(presentation, true)[0]?.rev;
}

describe("(a) local mode, the browser store wiped", () => {
  it("reads In this deck, data from <date time>, for every link with a record, and Update all paints nothing", async () => {
    await inserted(await seedText("EUR 3.2m"));
    await inserted(await seedText("EUR 4.8m"));
    const transport = await transportIn("local");

    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["inDeck", "inDeck"]);
    expect(await table(rows)).toEqual([
      { badge: "In this deck", meta: MOMENT_WORDS, updated: "just now" },
      { badge: "In this deck", meta: MOMENT_WORDS, updated: "just now" },
    ]);
    // The store alone still says what it said before this slice.
    const bare = await links.listLinks(transport.active());
    expect(bare.map((row) => row.status)).toEqual(["notPasted", "notPasted"]);

    const host = await quietHost();
    const before = customXmlSnapshot(presentation);
    const summary = await links.updateLinks(rows, transport.relay(), host);
    expect(summary).toMatchObject({ updated: 0, inDeck: 2, notPasted: 0 });
    expect(links.summarize(summary)).toBe("2 up to date with this deck");
    expect(host.refreshLink).not.toHaveBeenCalled();
    expect(host.refreshLinks).not.toHaveBeenCalled();
    expect(customXmlSnapshot(presentation)).toEqual(before);
  });

  it("is told by the deck's records as the pane read them back, not by this session's writes", async () => {
    await inserted(await seedText("EUR 3.2m"));
    const boot = await import("../src/ppt/deck-boot");
    await boot.loadDeckStore();
    const transport = await transportIn("local");
    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["inDeck"]);
  });
});

describe("(b) relay mode, a link past the relay's 30 days", () => {
  it("reads In this deck, not Source missing", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    relay.links.delete(item.id);
    const transport = await transportIn("relay");

    expect((await links.listLinks(relay)).map((row) => row.status)).toEqual([
      "missing",
    ]);
    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["inDeck"]);
    const summary = await links.updateLinks(rows, transport.relay());
    expect(summary).toMatchObject({ inDeck: 1, missing: 0, updated: 0 });
  });
});

describe("(c) the transport holds a newer rev", () => {
  it("reads Update available and Update all paints the newer one, the record following", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    await pushText(item, "EUR 3.9m");
    const transport = await transportIn("relay");

    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["updateAvailable"]);
    const summary = await links.updateLinks(rows, transport.relay());
    expect(summary).toMatchObject({ updated: 1, inDeck: 0, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.text).toBe("EUR 3.9m");
    expect(recordRev()).toBe(2);
    expect(recordsIn(presentation, true)[0]!.payload).toMatchObject({
      text: "EUR 3.9m",
    });
    const after = await links.listLinks(transport.relay());
    expect(after.map((row) => row.status)).toEqual(["current"]);
  });

  it("does the same in local mode, where a pasted copy newer than the slide wins over the deck's", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    const keys = await deriveLinkKeys(item.token);
    const store = await LocalStore.open(memoryPersistence());
    await store.ingest(
      {
        links: [
          {
            id: item.id,
            rev: LOCAL_REV_BASE + 5,
            sentAt: 1,
            blob: await seal(
              keys.enc,
              item.id,
              encodePayload(
                textPayload(
                  "EUR 4.4m",
                  new Date(Date.now() + 1000).toISOString(),
                ),
              ),
            ),
          },
        ],
        inbox: [],
      },
      new Set([item.id]),
    );
    const { layeredRelay } = await import("../src/ppt/layered-relay");
    const { deckRecords } = await import("../src/ppt/deck-boot");
    const layered = layeredRelay(store, deckRecords);

    const rows = await links.listLinks(layered);
    expect(rows.map((row) => row.status)).toEqual(["updateAvailable"]);
    await links.updateLinks(rows, layered);
    expect(presentation.slides[0]!.shapes[0]!.text).toBe("EUR 4.4m");
    expect(recordRev()).toBe(LOCAL_REV_BASE + 5);
  });
});

describe("(d) a record that holds no data", () => {
  it("keeps the old words: a picture link is Not pasted yet or Source missing, never In this deck", async () => {
    const item = await inserted(await seedLink(fakePng(800, 400)));
    expect(recordsIn(presentation, true)[0]).toMatchObject({
      dataKept: false,
      payload: null,
    });
    relay.links.delete(item.id);

    const relayRows = await links.listLinks(
      (await transportIn("relay")).relay(),
    );
    expect(relayRows.map((row) => row.status)).toEqual(["missing"]);
    const localRows = await links.listLinks(
      (await transportIn("local")).relay(),
    );
    expect(localRows.map((row) => row.status)).toEqual(["notPasted"]);
    const summary = await links.updateLinks(
      localRows,
      (await transportIn("local")).relay(),
    );
    expect(summary).toMatchObject({ notPasted: 1, inDeck: 0 });
  });
});

describe("(e) the Sent column", () => {
  it("shows the tag's time when the transport has none, never 'never'", async () => {
    const item = await inserted(await seedLink(fakePng(800, 400)));
    relay.links.delete(item.id);
    vi.setSystemTime(new Date(MOMENT.getTime() + 3 * 24 * 3600 * 1000));
    const rows = await links.listLinks((await transportIn("local")).relay());

    expect(rows[0]!.pushedAt).toBeNull();
    const tagTime = JSON.parse(
      presentation.slides[0]!.shapes[0]!.tags.get(TAG_LINK)!,
    ) as { pushedAt: string };
    expect((await table(rows))[0]!.updated).toBe(
      relativeTime(Date.parse(tagTime.pushedAt) / 1000),
    );
    expect((await table(rows))[0]!.updated).toBe("3 days ago");
  });
});

describe("(f) Revert on a row the deck answered", () => {
  it("refuses in words and writes nothing", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    await pushText(item, "EUR 3.9m");
    await links.updateLinks(await links.listLinks(relay), relay);
    expect(recordRev()).toBe(2);
    relay.links.delete(item.id);
    const transport = await transportIn("relay");
    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["inDeck"]);

    const { revertLinks, DECK_KEEPS_ONE } = await import("../src/ppt/revert");
    const host = await quietHost();
    const before = customXmlSnapshot(presentation);
    const summary = await revertLinks(rows, transport.relay(), host);
    expect(summary.reverted).toBe(0);
    expect(summary.failures).toEqual([
      `revert Model!B4:F12 text: ${DECK_KEEPS_ONE}`,
    ]);
    expect(DECK_KEEPS_ONE).toContain("keeps only the version on the slide");
    expect(host.refreshLink).not.toHaveBeenCalled();
    expect(host.refreshLinks).not.toHaveBeenCalled();
    expect(presentation.slides[0]!.shapes[0]!.text).toBe("EUR 3.9m");
    expect(customXmlSnapshot(presentation)).toEqual(before);
  });
});

describe("(g) what the floor never covers for", () => {
  it("a wrong key keeps its own word, record or not", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    relay.links.get(item.id)!.auth = "someone-else";
    const transport = await transportIn("relay");
    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["wrongKey"]);
    const summary = await links.updateLinks(rows, transport.relay());
    expect(summary).toMatchObject({ wrongKey: 1, inDeck: 0 });
  });

  it("a relay fault keeps its own sentence, on the list and on Update all", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    relay.links.delete(item.id);
    const transport = await transportIn("relay");
    const fault = new RelayError("server", "relay: 503", 503);
    vi.spyOn(relay, "status").mockRejectedValue(fault);
    await expect(links.listLinks(transport.relay())).rejects.toBe(fault);

    const rows = await links.listLinks((await transportIn("local")).relay());
    vi.spyOn(relay, "fetchLinks").mockRejectedValue(fault);
    vi.spyOn(relay, "getLink").mockRejectedValue(fault);
    const summary = await links.updateLinks(rows, transport.relay());
    expect(summary).toMatchObject({ failed: 1, inDeck: 0 });
    expect(summary.failures).toEqual([
      `Model!B4:F12 text: ${relayReason(fault)!}`,
    ]);
  });
});

describe("(i) a record at another rev than the slide's", () => {
  it("is not the slide's state: the transport's old words stay and nothing is claimed", async () => {
    const item = await inserted(await seedText("EUR 3.2m"));
    relay.links.delete(item.id);
    const record = recordsIn(presentation, true)[0]!;
    await removeAllParts();
    await seedParts([encodeDeckRecord({ ...record, rev: record.rev + 1 })]);
    await (await import("../src/ppt/deck-boot")).loadDeckStore();

    const local = await transportIn("local");
    const localRows = await links.listLinks(local.relay());
    expect(localRows.map((row) => row.status)).toEqual(["notPasted"]);
    const summary = await links.updateLinks(localRows, local.relay());
    expect(summary).toMatchObject({ notPasted: 1, inDeck: 0, updated: 0 });

    const remote = await transportIn("relay");
    const rows = await links.listLinks(remote.relay());
    expect(rows.map((row) => row.status)).toEqual(["missing"]);
  });
});

describe("(h) a Not on a slide row", () => {
  it("keeps its own words beside a row the deck answered", async () => {
    await inserted(await seedText("EUR 3.2m"));
    await inserted(await seedText("EUR 4.8m"));
    presentation.deleteShape(presentation.slides[0]!.shapes[0]!.id);
    const transport = await transportIn("local");

    const rows = await links.listLinks(transport.relay());
    expect(rows.map((row) => row.status)).toEqual(["inDeck"]);
    const { deckRecords } = await import("../src/ppt/deck-boot");
    const { drawingFrom } = await import("../src/ppt/deck-drawable");
    const { offSlideRows } = await import("../src/ppt/deck-rows");
    const off = offSlideRows(
      rows.map((row) => row.found),
      deckRecords(),
      true,
      drawingFrom,
    );
    expect(off).toHaveLength(1);
    const { toRowViews } = await import("../src/ppt/actions");
    const { renderLinkRows } = await import("../src/ppt/views");
    const body = document.createElement("tbody");
    renderLinkRows(body, toRowViews(rows, new Set()), () => undefined, {
      rows: off,
      onPutBack: () => undefined,
      onRemove: () => undefined,
    });
    expect(
      [...body.querySelectorAll(".badge")].map((badge) => badge.textContent),
    ).toEqual(["In this deck", "Not on a slide"]);
  });
});
