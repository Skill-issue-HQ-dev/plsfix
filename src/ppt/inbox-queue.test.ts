import { beforeEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../link/model";
import type * as QueueModule from "./inbox-queue";

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

const WEEK = 7 * 24 * 3600 * 1000;

// `at` is the export time Excel stamped on the link (an item's own createdAt),
// which a re-announce carries unchanged.
function item(id: string, at: number): InboxItem {
  return {
    id,
    token: `token-${id}`,
    kind: "range",
    label: `Model!B4:F12 (${id})`,
    src: SRC,
    createdAt: new Date(at).toISOString(),
  };
}

// The usual row: posted when the link was exported, so the relay's time and
// the item's own time agree.
function row(
  id: string,
  createdAt: number,
): { item: InboxItem; createdAt: number } {
  return { item: item(id, createdAt), createdAt };
}

// The pasted set lives for the pane's lifetime, so every test gets a new one.
let queue: typeof QueueModule;

beforeEach(async () => {
  vi.resetModules();
  queue = await import("./inbox-queue");
});

describe("queueOrder", () => {
  it("puts the newest export first whatever order the relay sent", () => {
    const ordered = queue.queueOrder([
      row("older", 1_000),
      row("newest", 3_000),
      row("middle", 2_000),
    ]);
    expect(ordered.map((one) => one.id)).toEqual(["newest", "middle", "older"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("newest");
  });

  it("drops an export this pane already pasted, however often it is asked", () => {
    queue.rememberPasted("pasted");

    const ordered = queue.queueOrder([
      row("pasted", 3_000),
      row("fresh", 1_000),
    ]);

    expect(ordered.map((one) => one.id)).toEqual(["fresh"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("fresh");
  });

  it("answers nothing when every waiting row is one this pane pasted", () => {
    queue.rememberPasted("a");
    expect(queue.queueOrder([row("a", 1_000)])).toEqual([]);
    expect(queue.latestInboxItem([])).toBeNull();
  });
});

// Excel's Push re-announces every link it pushes, so the relay's Inbox can
// hold a row for a link the deck already has on a slide. The deck's own tags
// are the authority on what to offer: a row it holds is hidden, a row whose
// object was deleted is welcome again.
describe("queueOrder against what the deck holds", () => {
  it("hides a row whose link the deck already holds", () => {
    const ordered = queue.queueOrder(
      [row("held", 3_000), row("free", 1_000)],
      new Set(["held"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["free"]);
  });

  it("keeps the newest-first order of the rows it shows, and the latest is the newest shown", () => {
    const ordered = queue.queueOrder(
      [row("old", 1_000), row("held", 5_000), row("new", 3_000)],
      new Set(["held"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["new", "old"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("new");
  });

  it("shows a row for an object the deck no longer holds, even one this pane pasted earlier", () => {
    queue.rememberPasted("deleted");

    const ordered = queue.queueOrder([row("deleted", 2_000)], new Set());

    expect(ordered.map((one) => one.id)).toEqual(["deleted"]);
  });

  it("still hides a pasted export while the deck holds it", () => {
    queue.rememberPasted("kept");
    const ordered = queue.queueOrder(
      [row("kept", 2_000), row("fresh", 1_000)],
      new Set(["kept"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["fresh"]);
  });

  it("forgets a pasted export the deck has lost, so an unreadable deck later does not hide its row", () => {
    queue.rememberPasted("deleted");
    queue.queueOrder([], new Set());

    // No deck read this time: only this pane's own memory is left to go by.
    const ordered = queue.queueOrder([row("deleted", 2_000)]);

    expect(ordered.map((one) => one.id)).toEqual(["deleted"]);
  });

  it("goes by this pane's own memory when the deck could not be read", () => {
    queue.rememberPasted("pasted");
    const ordered = queue.queueOrder([
      row("pasted", 3_000),
      row("fresh", 1_000),
    ]);
    expect(ordered.map((one) => one.id)).toEqual(["fresh"]);
  });
});

// The deck's own records are the deck too: the Linked objects list shows a
// link whose object was deleted as "Not on a slide" and puts it back from its
// row, so the Inbox leaves that link's row out. `recorded` is the links whose
// record the host can draw again - one it cannot (identity only, a chart this
// host draws as a picture) cannot put anything back, so its Inbox row stays.
describe("queueOrder against the deck's records", () => {
  it("hides a row the deck holds a record for, though no shape carries it", () => {
    const ordered = queue.queueOrder(
      [row("deleted", 3_000), row("free", 1_000)],
      new Set(),
      new Set(["deleted"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["free"]);
  });

  it("hides it when the deck's tags could not be read at all", () => {
    const ordered = queue.queueOrder(
      [row("deleted", 3_000), row("free", 1_000)],
      undefined,
      new Set(["deleted"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["free"]);
  });

  it("hides what the tags name and what the records name, each on its own", () => {
    const ordered = queue.queueOrder(
      [row("tagged", 3_000), row("recorded", 2_000), row("free", 1_000)],
      new Set(["tagged"]),
      new Set(["recorded"]),
    );
    expect(ordered.map((one) => one.id)).toEqual(["free"]);
    expect(queue.latestInboxItem(ordered)?.id).toBe("free");
  });

  it("changes nothing when no record is named", () => {
    const rows = [row("a", 2_000), row("b", 1_000)];
    expect(queue.queueOrder(rows, new Set(), new Set())).toEqual(
      queue.queueOrder(rows, new Set()),
    );
  });

  it("leaves this pane's memory of what it pasted as the tags left it", () => {
    queue.rememberPasted("deleted");
    queue.queueOrder([], new Set(), new Set(["deleted"]));
    // The tags no longer name the object, so the memory of its paste goes; the
    // record alone keeps the row hidden, and only while the record is there.
    const ordered = queue.queueOrder([row("deleted", 2_000)]);
    expect(ordered.map((one) => one.id)).toEqual(["deleted"]);
  });
});

describe("shownInbox", () => {
  const scanned =
    (...ids: string[]) =>
    async () =>
      ids.map((id) => ({ tag: { id } }));

  it("reads the deck and hides what it holds", async () => {
    const shown = await queue.shownInbox(
      [row("held", 3_000), row("free", 1_000)],
      scanned("held"),
    );
    expect(shown.map((one) => one.id)).toEqual(["free"]);
  });

  it("hides the rows the deck's records name as well as the ones its tags name", async () => {
    const shown = await queue.shownInbox(
      [row("tagged", 4_000), row("recorded", 3_000), row("free", 1_000)],
      scanned("tagged"),
      new Set(["recorded"]),
    );
    expect(shown.map((one) => one.id)).toEqual(["free"]);
  });

  it("still hides a recorded row when the deck cannot be read", async () => {
    const shown = await queue.shownInbox(
      [row("recorded", 3_000), row("other", 1_000)],
      async () => {
        throw new Error("PowerPoint stopped answering");
      },
      new Set(["recorded"]),
    );
    expect(shown.map((one) => one.id)).toEqual(["other"]);
  });

  it("does not read the deck when no row waits", async () => {
    let reads = 0;
    const shown = await queue.shownInbox([], async () => {
      reads += 1;
      return [];
    });
    expect(shown).toEqual([]);
    expect(reads).toBe(0);
  });

  it("shows every row, minus what this pane pasted, when the deck cannot be read", async () => {
    queue.rememberPasted("pasted");
    const shown = await queue.shownInbox(
      [row("pasted", 3_000), row("other", 1_000)],
      async () => {
        throw new Error("PowerPoint stopped answering");
      },
    );
    expect(shown.map((one) => one.id)).toEqual(["other"]);
  });
});

// The relay's row time moves on every Push (a re-post renews it), so the list
// is ordered by when a row was last posted. "Paste latest linked" means the
// newest EXPORT, which is the item's own createdAt: set once in Excel, carried
// unchanged by every re-announce.
describe("latestInboxItem", () => {
  it("is the newest export, not the row the relay posted last", () => {
    const list = [item("pushed-again", 1_000), item("fresh", 9_000_000)];
    expect(queue.latestInboxItem(list)?.id).toBe("fresh");
  });

  it("keeps a week-old link pushed 30 s after a fresh export second to it", () => {
    const exportedNow = Date.parse("2026-10-03T10:00:00Z");
    const ordered = queue.queueOrder([
      { item: item("fresh", exportedNow), createdAt: 1_000 },
      { item: item("old", exportedNow - WEEK), createdAt: 1_030 },
    ]);

    // The list keeps the relay's order, last posted first ...
    expect(ordered.map((one) => one.id)).toEqual(["old", "fresh"]);
    // ... and the one-click paste still takes the export.
    expect(queue.latestInboxItem(ordered)?.id).toBe("fresh");
  });

  it("keeps the list order between exports made at the same moment", () => {
    const list = [item("first", 5_000), item("second", 5_000)];
    expect(queue.latestInboxItem(list)?.id).toBe("first");
  });

  it("reads a createdAt it cannot parse as the oldest time there is", () => {
    const broken = { ...item("broken", 0), createdAt: "not a time" };
    expect(queue.latestInboxItem([broken, item("dated", 1_000)])?.id).toBe(
      "dated",
    );
    expect(queue.latestInboxItem([broken, broken])).toBe(broken);
  });

  it("answers nothing for an empty list", () => {
    expect(queue.latestInboxItem([])).toBeNull();
  });
});
