// src/link/wire.test.ts
// The batch fetch's body guard: the reasons a server may give, and only those.

import { describe, expect, it } from "vitest";
import { isFetchJsonBody } from "./wire";

const body = (reason: string): unknown => ({
  items: [],
  omitted: [{ id: "a", reason }],
});

describe("isFetchJsonBody", () => {
  it("believes every reason a relay or a local store gives", () => {
    for (const reason of [
      "unchanged",
      "missing",
      "auth",
      "deferred",
      "notPasted",
    ]) {
      expect(isFetchJsonBody(body(reason)), reason).toBe(true);
    }
  });

  it("does not believe a server that answers for the deck: inDeck is the pane's own word", () => {
    expect(isFetchJsonBody(body("inDeck"))).toBe(false);
  });
});
