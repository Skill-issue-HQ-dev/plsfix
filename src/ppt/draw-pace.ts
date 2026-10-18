// The pace of a chart draw: each chunk sized BEFORE its sync from the cost
// model (chart-cost.ts: the web charges for every add and for every shape
// already on the slide, and a batch past about 30 s is swallowed), the cap
// that halves when a sync comes back slower than the model said and doubles
// again when one comes back with room to spare, and the clock
// that times it.
// Invariant: only the pacing rule lives here; the numbers are chart-cost.ts's
// and the draw itself is chart-draw.ts.

import { systemClock, type Clock } from "../link/debounce";
import {
  chunkFor,
  SHAPES_PER_SYNC,
  SYNC_TARGET_SECONDS,
  syncSeconds,
  type DrawHost,
} from "./chart-cost";

export { SHAPES_PER_SYNC };

// A sync is slow when it took more than this many times its prediction AND
// longer than the target a sync is held to: a fast sync that was merely
// dearer than the model is not a reason to send fewer shapes.
export const SLOW_FACTOR = 1.5;

// Down on a slow sync, back up on a roomy one. The growth of a sync's cost with
// the shapes already on the slide is the model's (nextChunk); the cap answers
// what the model cannot see, a host slower than it was measured, and that is
// often a spell: on 05.10 PowerPoint for the web ran slow one afternoon and a
// cap that only halved drew a 92-shape chart one shape a sync, 29 syncs and
// 219 s where 11 syncs had taken 115 s that morning. A sync back in under half
// the target doubles the cap again, up to SHAPES_PER_SYNC (the model still
// bounds every chunk); the floor is one add a sync.
export const ROOMY_FACTOR = 0.5;

export function capAfter(
  cap: number,
  tookMs: number,
  predictedMs: number,
): number {
  const targetMs = SYNC_TARGET_SECONDS * 1000;
  if (tookMs > SLOW_FACTOR * predictedMs && tookMs > targetMs) {
    return Math.max(1, Math.floor(cap / 2));
  }
  const roomy = tookMs < ROOMY_FACTOR * targetMs;
  return roomy ? Math.min(SHAPES_PER_SYNC, cap * 2) : cap;
}

// What the next sync carries and how long the model says it will take: the
// model's chunk for the slide as it stands, under the cap a slow sync lowered
// and the shapes still to draw.
export function nextChunk(
  cap: number,
  population: number,
  remaining: number,
  host: DrawHost,
): { size: number; predictedMs: number } {
  const size = Math.max(
    1,
    Math.min(cap, chunkFor(population, host), remaining),
  );
  return { size, predictedMs: syncSeconds(size, population, host) * 1000 };
}

// The one clock a draw reads. The fake PowerPoint host charges a costed clock
// for every add (test/clock.ts), and a test points the drawer at the same one.
let clock: Clock = systemClock();

export function useDrawClock(next: Clock | null): void {
  clock = next ?? systemClock();
}

export function drawClock(): Clock {
  return clock;
}
