// The PowerPoint Inbox against a deck that already holds some of the links,
// end to end over the fake host and the fake relay. Excel's Push now posts an
// Inbox row for every link it pushes, so the relay lists rows for links that
// sit on a slide: those are hidden (the deck's tags decide), a row for an
// object the user deleted is offered again and puts the same link back, and a
// deck the host cannot read leaves the Inbox as it was. Strict load semantics.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { seal } from "../src/link/crypto";
import { encodeInboxItem, TAG_KEY, TAG_LINK } from "../src/link/model";
import type { InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
  type FakePptHelpers,
} from "./fakeppt";
import { bootPpt, memoryStore, pushAgain, seedLink } from "./ppt.support";
import { latestInboxItem } from "../src/ppt/inbox-queue";
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

// What Excel's Push does for one link: the same sealed row an export posts.
async function announce(item: InboxItem): Promise<void> {
  await relay.postInbox(
    ws.id,
    ws.auth,
    item.id,
    await seal(ws.enc, ws.id, encodeInboxItem(item)),
  );
}

async function waitingIds(): Promise<string[]> {
  return (await links.listInbox(ws, relay)).map((item) => item.id);
}

// The pane reopened over the same deck: its memory of what it pasted is gone,
// so only the deck's own tags can tell it what is already on a slide.
async function reopenPane(): Promise<void> {
  vi.resetModules();
  links = await import("../src/ppt/links");
}

describe("the Inbox against what the deck holds", () => {
  it("hides the row of a link the deck holds, though the relay still lists it", async () => {
    const item = await seedLink(fakePng(800, 400));
    await announce(item);
    await links.insertFromInbox(item, ws, relay);
    expect(await waitingIds()).toEqual([]);

    // Excel pushes the link again: the relay has its row back.
    await reopenPane();
    await announce(item);

    expect(relay.inbox.size).toBe(1);
    expect(await waitingIds()).toEqual([]);
  });

  it("offers the row again once the object is deleted, and inserting it puts the same link back", async () => {
    const item = await seedLink(fakePng(800, 400));
    await announce(item);
    await links.insertFromInbox(item, ws, relay);
    presentation.deleteShape(presentation.slides[0]!.shapes[0]!.id);

    // Same pane session as the paste above, so its memory of that paste must
    // not keep the re-announced row away.
    await announce(item);
    const waiting = await links.listInbox(ws, relay);
    expect(waiting.map((one) => one.id)).toEqual([item.id]);

    await links.insertFromInbox(waiting[0]!, ws, relay);
    const back = presentation.slides[0]!.shapes[0]!;
    expect(JSON.parse(back.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      rev: 1,
    });
    expect(back.tags.get(TAG_KEY)).toBe(item.token);
    expect(await waitingIds()).toEqual([]);

    // It refreshes exactly as before: Excel's next push reaches it.
    await pushAgain(item, fakePng(100, 50));
    const rows = await links.listLinks(relay);
    expect(rows.map((row) => row.status)).toEqual(["updateAvailable"]);
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
  });

  it("shows only the rows the deck lacks, newest first, and the latest of them is the newest shown", async () => {
    const held = await seedLink(fakePng(100, 50));
    await announce(held);
    await links.insertFromInbox(held, ws, relay);
    await reopenPane();

    relay.now += 1;
    const older = await seedLink(fakePng(200, 100));
    await announce(older);
    relay.now += 1;
    const newer = await seedLink(fakePng(300, 150));
    await announce(newer);
    relay.now += 1;
    // The newest row of all belongs to the link the deck already holds.
    await announce(held);

    const shown = await links.listInbox(ws, relay);

    expect(shown.map((one) => one.id)).toEqual([newer.id, older.id]);
    expect(latestInboxItem(shown)?.id).toBe(newer.id);
  });

  // The relay renews a re-posted row's time, so a week-old link Excel pushed
  // 30 s after a fresh export is listed ahead of it. "Paste latest linked"
  // still means the newest EXPORT: the item's own createdAt, set once in Excel
  // and carried unchanged by every re-announce.
  it("keeps Paste latest linked on the newest export when an old link is pushed after it", async () => {
    const weekAgo = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString();
    const old = { ...(await seedLink(fakePng(100, 50))), createdAt: weekAgo };
    relay.now = 1000;
    const fresh = await seedLink(fakePng(200, 100));
    await announce(fresh);
    relay.now = 1030;
    await announce(old);

    const shown = await links.listInbox(ws, relay);

    expect(shown.map((one) => one.id)).toEqual([old.id, fresh.id]);
    expect(latestInboxItem(shown)?.id).toBe(fresh.id);
  });

  it("reads the deck only when something waits", async () => {
    const before = helpers.syncCount();
    expect(await waitingIds()).toEqual([]);
    expect(helpers.syncCount()).toBe(before);

    await announce(await seedLink(fakePng(100, 50)));
    expect(await waitingIds()).toHaveLength(1);
    expect(helpers.syncCount()).toBeGreaterThan(before);
  });

  it("lists every row when the deck cannot be read, minus what this pane pasted", async () => {
    const pasted = await seedLink(fakePng(100, 50));
    await announce(pasted);
    await links.insertFromInbox(pasted, ws, relay);
    await announce(pasted);
    const other = await seedLink(fakePng(200, 100));
    await announce(other);

    const before = helpers.syncCount();
    helpers.failNextSync(new Error("PowerPoint stopped answering"));
    const shown = await links.listInbox(ws, relay);

    // The deck was asked, the host refused, and the Inbox still answered.
    expect(helpers.syncCount()).toBeGreaterThan(before);
    expect(shown.map((one) => one.id)).toEqual([other.id]);
  });
});
