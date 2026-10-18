// A virtual Clock for the debouncer: time only moves when a test says so, and
// the callbacks that fall due run synchronously inside advance(). Keeps the
// three-second auto-push window out of the suite's wall clock. costedClock is
// the same clock with a price on every shapes.add* of the fake PowerPoint host.

import type { Clock, Timer } from "../src/link/debounce";

export interface VirtualClock extends Clock {
  advance(ms: number): void;
  waiting(): number;
}

export function virtualClock(): VirtualClock {
  let now = 0;
  let scheduled: { at: number; run: () => void }[] = [];

  return {
    now: () => now,
    after(ms: number, run: () => void): Timer {
      const entry = { at: now + ms, run };
      scheduled.push(entry);
      return {
        cancel: () => {
          scheduled = scheduled.filter((other) => other !== entry);
        },
      };
    },
    advance(ms: number): void {
      now += ms;
      const due = scheduled.filter((entry) => entry.at <= now);
      scheduled = scheduled.filter((entry) => entry.at > now);
      for (const entry of due) entry.run();
    },
    waiting: () => scheduled.length,
  };
}

// The two terms the web's slowness is made of (lessons 30.08 and 27.09: twelve
// rectangles in 1.7 s on a clean slide, twenty-four in 15 s at 45 shapes): a
// flat price per add and one more per shape the slide already holds.
export interface AddCosts {
  msPerAdd: number;
  msPerExistingShape: number;
}

// The one thing the fake host needs of the clock it is given.
export interface AddCostClock {
  charge(existingShapes: number): void;
}

export interface CostedClock extends VirtualClock, AddCostClock {
  // Mutable, so a test can reprice the host between adds.
  costs: AddCosts;
  // How many adds the host has charged.
  charged(): number;
}

// A virtualClock the fake PowerPoint host charges (installFakePpt({ clock })):
// each shapes.add* advances it by msPerAdd + msPerExistingShape x the shapes
// on the slide before the add, so a drawer that times a chunk off this clock
// can be shown a slow host. Both prices default to 0: an unpriced clock is a
// virtualClock that counts adds.
export function costedClock(costs: Partial<AddCosts> = {}): CostedClock {
  const clock = virtualClock();
  const prices: AddCosts = { msPerAdd: 0, msPerExistingShape: 0, ...costs };
  let adds = 0;

  return {
    ...clock,
    costs: prices,
    charged: () => adds,
    charge(existingShapes: number): void {
      adds += 1;
      const ms = prices.msPerAdd + prices.msPerExistingShape * existingShapes;
      if (ms > 0) clock.advance(ms);
    },
  };
}
