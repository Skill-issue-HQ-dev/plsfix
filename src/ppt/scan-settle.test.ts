// When the pane may believe a scan: only after two consecutive scans, both taken
// SCAN_SETTLE_MS or more after the first and that far apart, show the same slides
// and links (about 30 s after boot): a far slide still loading reads as empty.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createScanSettle,
  createSettleTimer,
  SCAN_SETTLE_MS,
  SETTLE_RESCANS_MAX,
  SETTLE_RETRY_MS,
  type ScanSeen,
} from "./scan-settle";

const seen = (slides: number, ...ids: string[]): ScanSeen => ({
  slides,
  ids: new Set(ids),
});

describe("a scan settling", () => {
  let clock = 0;
  const now = (): number => clock;
  beforeEach(() => {
    clock = 1_000_000;
  });

  it("never settles on the first scan", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);
    clock += 10 * SCAN_SETTLE_MS;
    expect(settle.isSettled()).toBe(false);
  });

  it("does not settle on one late scan, however well it agrees with the scan at boot", () => {
    // A far slide that reads empty at boot and still at the first rescan would
    // read as a deleted object: one late scan is not two agreeing ones.
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);
  });

  it("does not settle on scans taken before SCAN_SETTLE_MS, however many and however well they agree", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    for (let step = 0; step < 5; step += 1) {
      clock += (SCAN_SETTLE_MS - 1) / 6;
      settle.observe(seen(3, "a"));
    }
    expect(settle.isSettled()).toBe(false);
  });

  it("settles on two late scans SCAN_SETTLE_MS apart that show the same slides and links", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a", "b"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a", "b"));
    expect(settle.isSettled()).toBe(false);
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a", "b"));
    expect(settle.isSettled()).toBe(true);
  });

  it("does not settle on two late scans closer than SCAN_SETTLE_MS apart", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS - 1;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);
    // The one after that is a full interval after the one before it.
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(true);
  });

  it("does not settle while the late scans disagree, and settles on the next pair that agrees", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(40, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(40, "a"));
    clock += SCAN_SETTLE_MS;
    // The deck is still loading: this scan shows a link the one before lacked.
    settle.observe(seen(40, "a", "far"));
    expect(settle.isSettled()).toBe(false);
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(40, "a", "far"));
    expect(settle.isSettled()).toBe(true);
  });

  it("does not settle on late scans where a link went away either: only the same links agree", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a", "b"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a", "b"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(true);
  });

  it("does not settle while the slide count moves", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(4, "a"));
    expect(settle.isSettled()).toBe(false);
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(4, "a"));
    expect(settle.isSettled()).toBe(true);
  });

  it("counts the time from the first scan, not from when the settle was made", () => {
    const settle = createScanSettle(now);
    clock += 60_000;
    settle.observe(seen(3, "a"));
    clock += 1_000;
    settle.observe(seen(3, "a"));
    clock += 1_000;
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);
  });

  it("stays settled: a later scan, even one that shows more, takes nothing back", () => {
    const settle = createScanSettle(now);
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    clock += SCAN_SETTLE_MS;
    settle.observe(seen(3, "a"));
    settle.observe(seen(5, "a", "b", "c"));
    expect(settle.isSettled()).toBe(true);
  });
});

describe("the clock a settle reads", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("is the monotonic one: the wall clock jumping settles nothing", () => {
    const settle = createScanSettle();
    settle.observe(seen(3, "a"));
    // A sleep, an NTP step: the date moves by minutes, the time since boot not.
    vi.setSystemTime(Date.now() + 10 * 60_000);
    settle.observe(seen(3, "a"));
    vi.setSystemTime(Date.now() + SCAN_SETTLE_MS);
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(false);

    // Time that really passed counts.
    vi.advanceTimersByTime(SCAN_SETTLE_MS);
    settle.observe(seen(3, "a"));
    vi.advanceTimersByTime(SCAN_SETTLE_MS);
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(true);
  });
});

describe("the one rescan after boot", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  // A pane stand-in: every rescan feeds the settle the next scan in `scans`.
  function pane(scans: ScanSeen[], extra: { busy?: () => boolean } = {}) {
    const settle = createScanSettle();
    let rescans = 0;
    const alive = { value: true };
    const timer = createSettleTimer({
      settle,
      rescan: async () => {
        rescans += 1;
        settle.observe(scans[Math.min(rescans, scans.length - 1)]!);
      },
      busy: extra.busy ?? (() => false),
      alive: () => alive.value,
    });
    settle.observe(scans[0]!);
    return { settle, timer, rescans: () => rescans, alive };
  }

  it("rescans SCAN_SETTLE_MS after it is armed and again SCAN_SETTLE_MS later, and settles on the second", async () => {
    const { settle, timer, rescans } = pane([
      seen(3, "a"),
      seen(3, "a"),
      seen(3, "a"),
    ]);
    timer.arm();
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS - 1);
    expect(rescans()).toBe(0);
    await vi.advanceTimersByTimeAsync(1);
    // One late scan agreeing with the boot scan settles nothing.
    expect(rescans()).toBe(1);
    expect(settle.isSettled()).toBe(false);
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(rescans()).toBe(2);
    expect(settle.isSettled()).toBe(true);
    // Settled: no timer is left to fire.
    await vi.advanceTimersByTimeAsync(10 * SCAN_SETTLE_MS);
    expect(rescans()).toBe(2);
  });

  it("rescans again when the second rescan still found more than the first", async () => {
    const scans = [
      seen(40, "a"),
      seen(40, "a"),
      seen(40, "a", "far"),
      seen(40, "a", "far"),
    ];
    const { settle, timer, rescans } = pane(scans);
    timer.arm();
    await vi.advanceTimersByTimeAsync(2 * SCAN_SETTLE_MS);
    expect(settle.isSettled()).toBe(false);
    expect(rescans()).toBe(2);
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(rescans()).toBe(3);
    expect(settle.isSettled()).toBe(true);
  });

  it("gives up after SETTLE_RESCANS_MAX rescans of a deck that keeps growing", async () => {
    const growing = Array.from({ length: SETTLE_RESCANS_MAX + 3 }, (_, i) =>
      seen(9, ...Array.from({ length: i + 1 }, (_n, k) => `id${String(k)}`)),
    );
    const { settle, timer, rescans } = pane(growing);
    timer.arm();
    await vi.advanceTimersByTimeAsync(20 * SCAN_SETTLE_MS);
    expect(rescans()).toBe(SETTLE_RESCANS_MAX);
    expect(settle.isSettled()).toBe(false);
  });

  it("does not rescan when the user's own refreshes settled the scan first", async () => {
    const settle = createScanSettle();
    let rescans = 0;
    const timer = createSettleTimer({
      settle,
      rescan: async () => {
        rescans += 1;
      },
      busy: () => false,
      alive: () => true,
    });
    settle.observe(seen(3, "a"));
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS + 1_000);
    // The user's first late scan, then the timer armed 4 s later: it fires at
    // 35 s, after the user's second late scan at 31 s.
    settle.observe(seen(3, "a"));
    await vi.advanceTimersByTimeAsync(4_000);
    timer.arm();
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS - 4_000);
    settle.observe(seen(3, "a"));
    expect(settle.isSettled()).toBe(true);
    await vi.advanceTimersByTimeAsync(2 * SCAN_SETTLE_MS);
    expect(rescans).toBe(0);
    expect(vi.getTimerCount()).toBe(0);
  });

  it("waits while an action is running, and asks again shortly", async () => {
    let busy = true;
    const { settle, timer, rescans } = pane(
      [seen(3, "a"), seen(3, "a"), seen(3, "a")],
      { busy: () => busy },
    );
    timer.arm();
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(rescans()).toBe(0);
    busy = false;
    await vi.advanceTimersByTimeAsync(SETTLE_RETRY_MS);
    expect(rescans()).toBe(1);
    expect(settle.isSettled()).toBe(false);
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(rescans()).toBe(2);
    expect(settle.isSettled()).toBe(true);
  });

  it("waits for an action however long it runs, and rescans once the pane is idle", async () => {
    let busy = true;
    const { timer, rescans } = pane(
      [seen(3, "a"), seen(3, "a"), seen(3, "a")],
      { busy: () => busy },
    );
    timer.arm();
    // Five minutes of a chart insert: far past the 20 s it used to wait.
    await vi.advanceTimersByTimeAsync(5 * 60_000);
    expect(rescans()).toBe(0);
    busy = false;
    await vi.advanceTimersByTimeAsync(SETTLE_RETRY_MS);
    expect(rescans()).toBe(1);
  });

  it("is cleared by stop, and by a pane that is no longer there", async () => {
    const stopped = pane([seen(3, "a"), seen(3, "a")]);
    stopped.timer.arm();
    stopped.timer.stop();
    await vi.advanceTimersByTimeAsync(5 * SCAN_SETTLE_MS);
    expect(stopped.rescans()).toBe(0);
    expect(vi.getTimerCount()).toBe(0);

    const gone = pane([seen(3, "a"), seen(3, "a")]);
    gone.timer.arm();
    gone.alive.value = false;
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(gone.rescans()).toBe(0);
    expect(vi.getTimerCount()).toBe(0);
  });

  it("holds one timer at most, however often it is armed", async () => {
    const { timer } = pane([seen(3, "a"), seen(3, "a")]);
    timer.arm();
    timer.arm();
    timer.arm();
    expect(vi.getTimerCount()).toBe(1);
  });

  it("survives a rescan that fails", async () => {
    const settle = createScanSettle();
    settle.observe(seen(3, "a"));
    let calls = 0;
    const timer = createSettleTimer({
      settle,
      rescan: async () => {
        calls += 1;
        throw new Error("PowerPoint stopped answering");
      },
      busy: () => false,
      alive: () => true,
    });
    timer.arm();
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    expect(calls).toBe(1);
    expect(settle.isSettled()).toBe(false);
  });
});
