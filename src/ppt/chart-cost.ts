// What drawing a chart costs on this kind of host, in seconds, and every rule
// built on it: how many adds one sync carries, what the whole draw is expected
// to take, the line above which the insert asks first, and when no draw can
// work at all. Pure arithmetic from S1's measurements (docs/superpowers/specs/
// 2026-10-03-s1-spike.md, 1.2 and 7): a sync of n adds onto a slide that holds
// p shapes takes c + a n + b (n p + n (n - 1) / 2) seconds.
// Invariant: every number lives in this file; the desktop row is NOT measured.

export type DrawHost = "web" | "desktop";

// The most adds one sync carries, however light the slide is.
export const SHAPES_PER_SYNC = 12;
// The length a sync is held to: 2.5 times under the 30 s above which no web
// batch ever answered (S1 1.3), so a chunk shrinks as the slide fills.
export const SYNC_TARGET_SECONDS = 12;
// A single add predicted past this is not offered a draw at all: every web
// batch modelled under 20 s answered (48 of 48), the next band lost one in ten.
export const SAFE_SYNC_SECONDS = 20;
// Above this many estimated seconds the insert asks "Draw it / Insert as
// picture" instead of starting a draw nobody expected to take minutes.
export const ASK_ABOVE_SECONDS = 120;

// How long the pane waits for one round trip before it calls the host silent.
// The web: the longest add sync that ever answered took 27.1 s and none
// predicted over 30 s did (S1 7), so 35 s; a longer wait only adds dead time to
// a swallow. The desktop has never been slow enough to time: kept at 60 s.
export const SYNC_TIMEOUT_MS: Readonly<Record<DrawHost, number>> = {
  web: 35_000,
  desktop: 60_000,
};

export interface HostCost {
  // Seconds a sync costs before it adds anything (c).
  perSync: number;
  // Seconds one add costs on an empty slide (a).
  perAdd: number;
  // Seconds more for every shape already on the slide when it is added (b).
  perExisting: number;
  // Seconds the grouping sync takes: addGroup, name and both tags.
  grouping: number;
}

// The web row is S1's mixed rectangles-and-labels fit (c 0.6, a 0.17, b 0.0205;
// rectangles 0.67 / 0.155 / 0.0190, text boxes 0.53 / 0.184 / 0.0225), addGroup
// measured at 0.84 to 2.87 s. The desktop row is a guess, NOT MEASURED: nothing
// on the desktop has ever been slow enough to time, so it waits for the Mac pass.
export const HOST_COSTS: Readonly<Record<DrawHost, HostCost>> = {
  web: { perSync: 0.6, perAdd: 0.17, perExisting: 0.0205, grouping: 2 },
  desktop: { perSync: 0.05, perAdd: 0.04, perExisting: 0.0002, grouping: 0.5 },
};

// One sync of `adds` shapes onto a slide holding `population`, group children
// included: the i-th add (from 0) is priced with population + i shapes there,
// which sums to adds x population + adds x (adds - 1) / 2 shape-waits.
export function syncSeconds(
  adds: number,
  population: number,
  host: DrawHost,
): number {
  const cost = HOST_COSTS[host];
  return (
    cost.perSync +
    cost.perAdd * adds +
    cost.perExisting * (adds * population + (adds * (adds - 1)) / 2)
  );
}

// The biggest chunk of at most SHAPES_PER_SYNC whose sync is predicted to stay
// inside the target on a slide of this population; never less than one.
export function chunkFor(population: number, host: DrawHost): number {
  for (let adds = SHAPES_PER_SYNC; adds > 1; adds -= 1) {
    if (syncSeconds(adds, population, host) <= SYNC_TARGET_SECONDS) {
      return adds;
    }
  }
  return 1;
}

// The whole draw: the syncs chunkFor plans from the slide as it stands, each
// priced at the population it will meet, plus the grouping. A nonsense count
// counts as none.
export function estimateSeconds(
  shapeCount: number,
  existingShapesOnSlide: number,
  host: DrawHost,
): number {
  const total = Math.max(0, Math.floor(shapeCount));
  const existing = Math.max(0, existingShapesOnSlide);
  if (total === 0) return 0;
  let seconds = HOST_COSTS[host].grouping;
  for (let done = 0; done < total;) {
    const adds = Math.min(chunkFor(existing + done, host), total - done);
    seconds += syncSeconds(adds, existing + done, host);
    done += adds;
  }
  return seconds;
}

// False once even a single add onto this slide is predicted past the safe
// line: no chunking helps, so the draw is not offered (about 940 shapes on the
// web).
export function canDraw(population: number, host: DrawHost): boolean {
  return syncSeconds(1, population, host) <= SAFE_SYNC_SECONDS;
}

// What the question says: whole minutes, and never fewer than the line it
// asks above.
export function askMinutes(seconds: number): number {
  return Math.max(Math.round(ASK_ABOVE_SECONDS / 60), Math.round(seconds / 60));
}
