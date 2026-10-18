// When the pane may believe a scan, and the one timer that asks again: a scan
// settles once two consecutive scans, both taken SCAN_SETTLE_MS or more after
// the first and SCAN_SETTLE_MS or more apart, agree. Invariant: no DOM, no
// Office.js, at most one timer, and nothing here throws.

import type { LinkId } from "../link/model";

// A deck that was just opened reads its far slides as empty for about 15 s (S1
// spike, section 6), so neither the first scan nor a scan that merely agrees
// with it can say an object is gone: a far slide still reading empty at both
// would be a deleted object. Two consecutive scans that are both late and a
// full interval apart, showing the same slides and the same links, can: about
// 30 s after boot at the earliest. SCAN_SETTLE_MS is also how long the timer
// waits between rescans.
export const SCAN_SETTLE_MS = 15_000;
// A deck that is still changing at the third rescan is left to the user's own
// refresh, which is a scan like any other.
export const SETTLE_RESCANS_MAX = 3;
// A rescan never starts under a running action; it asks again this soon.
export const SETTLE_RETRY_MS = 2_000;

// What one scan saw: how many slides the deck had and which links had a shape.
export interface ScanSeen {
  slides: number;
  ids: ReadonlySet<LinkId>;
}

export interface ScanSettle {
  observe(seen: ScanSeen): void;
  isSettled(): boolean;
}

function agrees(before: ScanSeen, after: ScanSeen): boolean {
  if (before.slides !== after.slides || before.ids.size !== after.ids.size) {
    return false;
  }
  for (const id of after.ids) if (!before.ids.has(id)) return false;
  return true;
}

// Fed every scan the pane takes, the boot one first. Settled for good once it
// is: a later scan that shows less is a real edit, not a loading deck. The
// clock is the monotonic one: a date set back or forward settles nothing.
export function createScanSettle(
  now: () => number = () => performance.now(),
): ScanSettle {
  let firstAt: number | null = null;
  // The last scan taken late enough to count: what the next one is held to.
  let late: { at: number; seen: ScanSeen } | null = null;
  let settled = false;
  return {
    observe(seen) {
      const at = now();
      firstAt ??= at;
      if (settled || at - firstAt < SCAN_SETTLE_MS) return;
      if (
        late !== null &&
        at - late.at >= SCAN_SETTLE_MS &&
        agrees(late.seen, seen)
      ) {
        settled = true;
      }
      late = { at, seen };
    },
    isSettled: () => settled,
  };
}

export interface SettleTimerDeps {
  settle: ScanSettle;
  // A scan the pane takes the way a refresh does, feeding `settle` itself.
  rescan: () => Promise<void>;
  busy: () => boolean;
  // False once the pane's page is gone (unloaded, or replaced).
  alive: () => boolean;
}

export interface SettleTimer {
  arm(): void;
  stop(): void;
}

// The rescan after boot: one timer, never two. It rescans silently when it
// fires, and again after each rescan that left the scan unsettled, up to
// SETTLE_RESCANS_MAX times; under a running action it waits, however long.
export function createSettleTimer(deps: SettleTimerDeps): SettleTimer {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let stopped = false;
  let rescans = 0;

  function clear(): void {
    if (timer !== null) clearTimeout(timer);
    timer = null;
  }
  function schedule(ms: number): void {
    clear();
    if (!stopped) timer = setTimeout(() => void tick(), ms);
  }
  async function tick(): Promise<void> {
    timer = null;
    if (stopped || !deps.alive() || deps.settle.isSettled()) return;
    if (deps.busy()) {
      // However long the action runs (a chart insert on the web takes minutes):
      // asked again as soon as it may be done, and never given up.
      schedule(SETTLE_RETRY_MS);
      return;
    }
    rescans += 1;
    try {
      await deps.rescan();
    } catch {
      // A rescan that failed is not a reason to believe the scan: the next
      // refresh the user makes is another look.
    }
    if (!deps.settle.isSettled() && rescans < SETTLE_RESCANS_MAX) {
      schedule(SCAN_SETTLE_MS);
    }
  }

  return {
    arm() {
      schedule(SCAN_SETTLE_MS);
    },
    stop() {
      stopped = true;
      clear();
    },
  };
}
