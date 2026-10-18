// Drives the fake clock until the work under test settles, so a per-sync
// deadline (SYNC_TIMEOUT_MS, 35 s on the web and 60 s elsewhere) is proven
// without a real minute; owns the one rule the "host never answers" suites
// share: a real event-loop turn per step.
// Invariant: a promise that never settles leaves after ~2 s and hits vitest's timeout.

import { setTimeout as realSetTimeout } from "node:timers";
import { vi } from "vitest";

import { SYNC_TIMEOUT_MS } from "../src/ppt/chart-draw";

const MAX_STEPS = 200;
const REAL_PAUSE_MS = 10;
// One step is the longer of the two deadlines, so a sync of either host has
// fired by the end of it.
const STEP_MS = Math.max(...Object.values(SYNC_TIMEOUT_MS));

/** Advance the fake clock past the longest sync deadline per step until `done()` holds. */
export async function settleUntil(done: () => boolean): Promise<void> {
  for (let step = 0; step < MAX_STEPS && !done(); step += 1) {
    await vi.advanceTimersByTimeAsync(STEP_MS);
    // A real turn: a decrypt or a fetch still in flight on a slow machine lands
    // here and only then arms the deadline timer the next step fires. Without
    // the pause the loop ran dry before the timer existed (CI runs, 12.09).
    await new Promise((resolve) => realSetTimeout(resolve, REAL_PAUSE_MS));
  }
}

/** The same loop for a promise in hand: until `work` settles, then its outcome. */
export async function settleHungSync<T>(work: Promise<T>): Promise<T> {
  let done = false;
  void work.then(
    () => (done = true),
    () => (done = true),
  );
  await settleUntil(() => done);
  return work;
}
