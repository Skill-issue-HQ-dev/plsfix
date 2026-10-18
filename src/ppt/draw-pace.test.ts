// The pace of a chart draw: each chunk sized before its sync from the cost
// model, the cap that halves when a sync comes back slower than that model
// said, and the clock that times it.

import { afterEach, describe, expect, it } from "vitest";
import { virtualClock } from "../../test/clock";
import {
  chunkFor,
  HOST_COSTS,
  SHAPES_PER_SYNC,
  SYNC_TARGET_SECONDS,
  syncSeconds,
} from "./chart-cost";
import {
  capAfter,
  drawClock,
  nextChunk,
  SLOW_FACTOR,
  useDrawClock,
} from "./draw-pace";

afterEach(() => {
  useDrawClock(null);
});

describe("the cap after a sync", () => {
  it("is 1.5 times the prediction and longer than the 12 s sync target that makes a sync slow", () => {
    expect(SLOW_FACTOR).toBe(1.5);
    expect(SYNC_TARGET_SECONDS).toBe(12);
  });

  it("halves when a sync took more than 1.5 times its prediction and more than 12 s", () => {
    expect(capAfter(SHAPES_PER_SYNC, 13_000, 8_000)).toBe(6);
    expect(capAfter(SHAPES_PER_SYNC, 40_000, 9_000)).toBe(6);
  });

  it("keeps the cap when the sync was only over 12 s, not over 1.5 times its prediction", () => {
    expect(capAfter(SHAPES_PER_SYNC, 20_000, 15_000)).toBe(12);
    expect(capAfter(SHAPES_PER_SYNC, 18_000, 12_000)).toBe(12);
  });

  it("keeps the cap when the sync was over 1.5 times its prediction but still under 12 s", () => {
    expect(capAfter(SHAPES_PER_SYNC, 11_000, 5_000)).toBe(12);
    expect(capAfter(SHAPES_PER_SYNC, 12_000, 4_000)).toBe(12);
  });

  it("keeps the cap for a sync that came back on time", () => {
    expect(capAfter(SHAPES_PER_SYNC, 0, 4_000)).toBe(12);
    expect(capAfter(SHAPES_PER_SYNC, 3_900, 4_000)).toBe(12);
  });

  it("halves again on the next slow sync, down to one and no further", () => {
    expect(capAfter(6, 30_000, 10_000)).toBe(3);
    expect(capAfter(3, 30_000, 10_000)).toBe(1);
    expect(capAfter(1, 30_000, 10_000)).toBe(1);
  });

  it("grows back, doubling, when a sync came back in under half the 12 s target, up to 12", () => {
    expect(capAfter(1, 4_000, 2_400)).toBe(2);
    expect(capAfter(3, 5_900, 3_000)).toBe(6);
    expect(capAfter(6, 100, 4_000)).toBe(12);
    expect(capAfter(SHAPES_PER_SYNC, 100, 4_000)).toBe(SHAPES_PER_SYNC);
  });

  it("keeps a lowered cap for a sync between half the target and the target", () => {
    expect(capAfter(3, 6_000, 5_000)).toBe(3);
    expect(capAfter(1, 11_000, 9_000)).toBe(1);
  });

  it("finds its way back after a slow spell: halved twice, then grown twice", () => {
    let cap = capAfter(SHAPES_PER_SYNC, 30_000, 10_000);
    cap = capAfter(cap, 30_000, 10_000);
    expect(cap).toBe(3);
    cap = capAfter(cap, 4_000, 3_000);
    cap = capAfter(cap, 5_000, 4_000);
    expect(cap).toBe(SHAPES_PER_SYNC);
  });
});

describe("the next chunk", () => {
  it("is what the model allows on the slide as it stands, under the cap and the shapes left", () => {
    expect(nextChunk(SHAPES_PER_SYNC, 0, 92, "web").size).toBe(12);
    expect(nextChunk(SHAPES_PER_SYNC, 36, 56, "web").size).toBe(11);
    expect(nextChunk(SHAPES_PER_SYNC, 100, 40, "web").size).toBe(5);
    expect(nextChunk(6, 0, 92, "web").size).toBe(6);
    expect(nextChunk(SHAPES_PER_SYNC, 0, 5, "web").size).toBe(5);
  });

  it("is never more than the model's own chunk, whatever the cap", () => {
    for (let population = 0; population <= 400; population += 7) {
      const { size } = nextChunk(SHAPES_PER_SYNC, population, 1_000, "web");
      expect(size).toBe(chunkFor(population, "web"));
    }
  });

  it("predicts the sync it sizes, in milliseconds, for the shapes it will really add", () => {
    const { size, predictedMs } = nextChunk(SHAPES_PER_SYNC, 24, 5, "web");
    expect(size).toBe(5);
    expect(predictedMs).toBeCloseTo(syncSeconds(5, 24, "web") * 1000, 6);
    expect(HOST_COSTS.web.perSync).toBeGreaterThan(0);
  });
});

describe("the draw clock", () => {
  it("reads the real clock until a test hands it another", () => {
    const real = drawClock();
    const before = Date.now();
    expect(real.now()).toBeGreaterThanOrEqual(before);
  });

  it("answers a handed clock, and the real one again after null", () => {
    const fake = virtualClock();
    useDrawClock(fake);
    fake.advance(1_234);
    expect(drawClock().now()).toBe(1_234);
    useDrawClock(null);
    expect(drawClock().now()).not.toBe(1_234);
  });
});
