// @vitest-environment jsdom
// The pane's rescans after a deck opens: the timer waits out a long action, a slow
// rescan never overwrites rows an action loaded since, a reload costs the scan's
// own round trips (the slide count rides its first sync).

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SCAN_SETTLE_MS, SETTLE_RETRY_MS } from "../src/ppt/scan-settle";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { openWithRecord } from "./ppt.deck-rows.open";
import { offSlideRows, refreshedAfterTheSettle } from "./ppt.deck-rows.support";
import {
  click,
  deck,
  hostHelpers,
  linkRows,
  paneRelay,
  settle,
} from "./stress.ppt.support";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

// One relay call the test opens by hand, so a pane flow can be caught mid-way.
function holdOnce(name: "status" | "listInbox"): () => void {
  const relay = paneRelay();
  let release = (): void => undefined;
  const gate = new Promise<void>((done) => {
    release = () => done();
  });
  const real = relay[name].bind(relay) as (...args: unknown[]) => unknown;
  let held = false;
  vi.spyOn(relay, name).mockImplementation(((...args: unknown[]) => {
    if (held) return real(...args);
    held = true;
    return gate.then(() => real(...args));
  }) as never);
  return release;
}

describe("a boot action longer than the timer used to wait", () => {
  beforeEach(() => {
    vi.useFakeTimers({ shouldAdvanceTime: true, advanceTimeDelta: 20 });
  });

  it("still gets its rows without a manual refresh: the timer waits for the pane to go idle", async () => {
    await openWithRecord();
    const release = holdOnce("listInbox");
    // An Inbox read that takes minutes: the pane is busy the whole time.
    click("refresh-inbox");
    // Not settle(): it waits for the latch this test holds shut.
    await vi.advanceTimersByTimeAsync(3 * 60_000);
    expect(offSlideRows()).toEqual([]);

    release();
    await settle();
    // Idle at last: the first rescan comes within a retry, the second a full
    // interval after it, and the row shows then.
    await vi.advanceTimersByTimeAsync(SETTLE_RETRY_MS + 500);
    await settle();
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS + 500);
    await settle();

    expect(offSlideRows()).toHaveLength(1);
  });
});

describe("a rescan that was slow", () => {
  beforeEach(() => {
    vi.useFakeTimers({ shouldAdvanceTime: true, advanceTimeDelta: 20 });
  });

  it("does not overwrite the rows an action loaded after it began", async () => {
    // The object is on slide 1 and its record is in the deck.
    await openWithRecord({ objectOn: 0 });
    expect(linkRows()).toHaveLength(1);
    const release = holdOnce("status");

    // The timer's rescan scans the deck, then waits on the relay's status.
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS + 500);
    await settle();
    // Meanwhile the user deletes the object and reads the list: it is gone.
    deck().deleteShape(deck().slides[0]!.shapes[0]!.id);
    click("refresh-links");
    await settle();
    expect(linkRows()).toHaveLength(0);

    release();
    await settle();

    // The rescan began before the deletion: its rows are stale, and dropped.
    expect(linkRows()).toHaveLength(0);
  });
});

describe("a reload", () => {
  beforeEach(() => {
    vi.useFakeTimers({ toFake: ["Date", "performance"] });
  });

  it("costs the store read and the scan and no round trip besides, before and after the settle", async () => {
    await openWithRecord({ objectOn: 0 });
    const spent = async (): Promise<number> => {
      const from = hostHelpers().syncCount();
      click("refresh-links");
      await settle();
      return hostHelpers().syncCount() - from;
    };
    // The part ids and the parts' text, then the slides, their shapes and the
    // shapes' tags: five round trips, the slide count among the scan's own.
    const STORE_AND_SCAN = 5;

    vi.advanceTimersByTime(SCAN_SETTLE_MS + 1_000);
    expect(await spent()).toBe(STORE_AND_SCAN);
    vi.advanceTimersByTime(SCAN_SETTLE_MS + 1_000);
    expect(await spent()).toBe(STORE_AND_SCAN);
    // Settled by now: the same again.
    await refreshedAfterTheSettle();
    expect(await spent()).toBe(STORE_AND_SCAN);
  });
});
