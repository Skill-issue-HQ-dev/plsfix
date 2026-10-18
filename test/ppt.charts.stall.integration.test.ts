// A chart's draw batch the host never answers (it neither applies nor rejects
// it, as PowerPoint for the web did on 08.09): the insert takes down what it
// drew, makes no picture, says nothing was inserted and leaves the Inbox row; a
// refresh keeps the old group and says so; the next chart draws normally. The
// web gives up on a silent sync at 35 s, the desktop at 60 s. Moved out of the
// audit suite at its 400-line cap. Strict load semantics are on.

import { setTimeout as realSetTimeout } from "node:timers";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  INSERT_STOPPED,
  REFRESH_STOPPED,
  type ChartData,
} from "../src/link/chart-model";
import { TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { InsertResult } from "../src/ppt/host";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { settleHungSync } from "./hung-sync";
import {
  bootPpt,
  columnChart,
  memoryStore,
  postToInbox,
  pushChart,
  seedChart,
} from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const COLUMN = columnChart(6);
// One title, six bars, six value labels, a baseline and six category labels.
const COLUMN_SHAPES = 20;
const LABEL = "Model: B4:F12";
// The insert's syncs on the desktop: 1 the selected slide, 2 the boxes already
// on it, 3 the first chunk of twelve, 4 the second chunk of eight, 5 the group.
// Both constants are how many syncs still answer before the one the host
// swallows (hangNextSync's own argument): three before the second chunk, four
// before the grouping sync.
const SECOND_CHUNK_SYNC = 3;
const GROUP_SYNC = 4;
// On the web one more sync comes first, the read of the slide's population.
const WEB_SECOND_CHUNK_SYNC = 4;

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

function box(one: FakePptShape) {
  return { left: one.left, top: one.top, width: one.width, height: one.height };
}

async function insert(
  data: ChartData = COLUMN,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

describe("a draw batch the host never answers", () => {
  // The Inbox row is what keeps the export alive: a swallowed insert leaves
  // it where it was, so the user presses Insert again.
  async function waitingChart(): Promise<{ ws: Workspace; item: InboxItem }> {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(COLUMN, PNG);
    await postToInbox(ws, item);
    return { ws, item };
  }

  async function inboxIds(ws: Workspace): Promise<string[]> {
    return (await links.listInbox(ws, relay)).map((item) => item.id);
  }

  it("takes down what it drew, inserts no picture and says nothing was inserted", async () => {
    const { ws, item } = await waitingChart();
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    vi.useFakeTimers();

    await expect(
      settleHungSync(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    // The twelve shapes the first chunk had already committed are gone, no
    // half-built group survived, and no picture took their place.
    expect(shapes()).toHaveLength(0);
    expect(await inboxIds(ws)).toEqual([item.id]);
  });

  it("says the same when the grouping sync is the swallowed one", async () => {
    const { ws, item } = await waitingChart();
    // The fifth sync groups, names and tags what the two chunks drew: the
    // host swallows that one, with all twenty shapes already on the slide.
    helpers.hangNextSync(GROUP_SYNC);
    vi.useFakeTimers();

    await expect(
      settleHungSync(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    expect(shapes()).toHaveLength(0);
    expect(await inboxIds(ws)).toEqual([item.id]);
  });

  it("fails rather than waits when the placement read is the swallowed one", async () => {
    const { ws, item } = await waitingChart();
    // The very first round trip - which slide is selected - never answers, so
    // there is nothing placed and nothing drawn: the insert has to end.
    helpers.hangNextSync();
    vi.useFakeTimers();

    await expect(
      settleHungSync(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    expect(shapes()).toHaveLength(0);
    expect(await inboxIds(ws)).toEqual([item.id]);
  });

  it("draws the next chart normally: one hung batch is not a jammed pane", async () => {
    const { ws, item: first } = await waitingChart();
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    vi.useFakeTimers();
    await expect(
      settleHungSync(links.insertFromInbox(first, ws, relay)),
    ).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    helpers.selectSlide(presentation.slides[1]!.id);
    const second = await seedChart(COLUMN, PNG);
    const placed = await links.insertFromInbox(second, ws, relay);

    const drawn = presentation.slides[1]!.shapes;
    expect(drawn).toHaveLength(1);
    expect(drawn[0]!.type).toBe("Group");
    expect(placed.note).toBeUndefined();
  });

  it("still fails the insert, with the same sentence, when the cleanup stops answering too", async () => {
    const { ws, item } = await waitingChart();
    // The draw's second chunk, then the cleanup's own first sync: a host that
    // has stopped answering altogether. Settling with a rejection is still
    // settling: the pane's guard releases the busy flag and the toast says
    // what happened.
    helpers.hangNextSync(SECOND_CHUNK_SYNC);
    helpers.hangNextSync(SECOND_CHUNK_SYNC + 1);
    vi.useFakeTimers();

    await expect(
      settleHungSync(links.insertFromInbox(item, ws, relay)),
    ).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    // No picture was put in its place, whatever the cleanup could not reach.
    expect(shapes().some((one) => one.fillImage !== null)).toBe(false);
    expect(await inboxIds(ws)).toEqual([item.id]);
  });

  it("keeps the old group where it was and says the old chart stays when a refresh stops answering", async () => {
    const { item } = await insert();
    const group = shapes()[0]!;
    group.left = 80;
    group.top = 320;
    group.width = 380;
    group.height = 190;
    await pushChart(item, columnChart(7), PNG);
    const rows = await links.listLinks(relay);

    // A group leaves the batched repaint before it spends a round trip, so
    // the update's own first sync is the draw's first chunk of twelve and its
    // second is the chunk of eleven that the host swallows.
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    const summary = await settleHungSync(links.updateLinks(rows, relay));
    vi.useRealTimers();

    // The chart did not update, so it is a failure, not an update with a note,
    // and the row says what the user is looking at.
    expect(summary).toMatchObject({ updated: 0, failed: 1 });
    expect(summary.failures).toEqual([`${LABEL}: ${REFRESH_STOPPED}`]);
    expect(summary.notes).toEqual([]);
    // The old group is untouched: same shape, same box, same rev, no picture.
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.id).toBe(group.id);
    expect(shapes()[0]!.type).toBe("Group");
    expect(shapes()[0]!.group!.shapes).toHaveLength(COLUMN_SHAPES);
    expect(box(shapes()[0]!)).toEqual({
      left: 80,
      top: 320,
      width: 380,
      height: 190,
    });
    expect(JSON.parse(shapes()[0]!.tags.get(TAG_LINK)!).rev).toBe(1);
  });

  it("redraws on the next Update all once the host answers again", async () => {
    const { item } = await insert();
    await pushChart(item, columnChart(7), PNG);
    const rows = await links.listLinks(relay);
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    await settleHungSync(links.updateLinks(rows, relay));
    vi.useRealTimers();

    const again = await links.updateLinks(await links.listLinks(relay), relay);

    expect(again).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).toBe("Group");
    expect(shapes()[0]!.group!.shapes).toHaveLength(23);
    expect(JSON.parse(shapes()[0]!.tags.get(TAG_LINK)!).rev).toBe(2);
  });
});

// How long a swallowed batch holds the pane: the fake clock advances a second
// at a time, with a real turn between steps for the pane's own async work, until
// the insert settles.
async function secondsUntilSettled(work: Promise<unknown>): Promise<number> {
  let done = false;
  void work.then(
    () => (done = true),
    () => (done = true),
  );
  let seconds = 0;
  while (!done && seconds < 120) {
    await vi.advanceTimersByTimeAsync(1_000);
    await new Promise((resolve) => realSetTimeout(resolve, 5));
    seconds += 1;
  }
  return seconds;
}

describe("how long the pane waits for a swallowed batch", () => {
  async function swallowedSecondChunk(
    platform: string,
    afterSyncs: number,
  ): Promise<{
    seconds: number;
    message: string;
    left: number;
    ws: Workspace;
    item: InboxItem;
  }> {
    helpers.setPlatform(platform);
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(COLUMN, PNG);
    await postToInbox(ws, item);
    helpers.hangNextSync(afterSyncs);
    vi.useFakeTimers();
    let message = "";
    const insertion = links
      .insertFromInbox(item, ws, relay)
      .catch((error: unknown) => {
        message = (error as Error).message;
      });
    const seconds = await secondsUntilSettled(insertion);
    vi.useRealTimers();
    return { seconds, message, left: shapes().length, ws, item };
  }

  it("is 35 s on PowerPoint for the web: the chunk is given up on, the shapes taken down", async () => {
    const run = await swallowedSecondChunk(
      "OfficeOnline",
      WEB_SECOND_CHUNK_SYNC,
    );

    expect(run.message).toBe(INSERT_STOPPED);
    expect(run.seconds).toBeGreaterThanOrEqual(35);
    expect(run.seconds).toBeLessThanOrEqual(37);
    expect(run.left).toBe(0);
    expect((await links.listInbox(run.ws, relay)).map((one) => one.id)).toEqual(
      [run.item.id],
    );
  });

  it("is 60 s on the desktop, which has never been slow enough to time", async () => {
    const run = await swallowedSecondChunk("PC", SECOND_CHUNK_SYNC);

    expect(run.message).toBe(INSERT_STOPPED);
    expect(run.seconds).toBeGreaterThanOrEqual(60);
    expect(run.seconds).toBeLessThanOrEqual(62);
    expect(run.left).toBe(0);
  });
});
