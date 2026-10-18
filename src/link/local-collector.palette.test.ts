// src/link/local-collector.palette.test.ts
// The brand colors a collector hands on: the bundle lists its fields one by one,
// so the palette travels only when the caller gave the collector one. An empty
// collector stays empty whatever palette it holds.
import { describe, expect, it } from "vitest";
import { LocalCollector } from "./local-collector";

const ID = "0123456789abcdef0123456789abcdef";
const PALETTE = ["#B27E54", "#282623", "#273631"];

describe("LocalCollector palette", () => {
  it("puts the palette it was given into the bundle", async () => {
    const collector = new LocalCollector(() => 1000, PALETTE);
    await collector.putLink(ID, "a", new Uint8Array([1]));
    expect(collector.bundle().palette).toEqual(PALETTE);
  });

  it("hands on a copy, not the caller's own list", async () => {
    const mine = [...PALETTE];
    const collector = new LocalCollector(() => 1000, mine);
    await collector.putLink(ID, "a", new Uint8Array([1]));
    mine.push("#111111");
    expect(collector.bundle().palette).toEqual(PALETTE);
  });

  it("has no palette key when it was given none", async () => {
    const collector = new LocalCollector(() => 1000);
    await collector.putLink(ID, "a", new Uint8Array([1]));
    expect(collector.bundle()).not.toHaveProperty("palette");
  });

  it("is still empty with a palette and nothing recorded", () => {
    const collector = new LocalCollector(() => 1000, PALETTE);
    expect(collector.isEmpty()).toBe(true);
  });
});
