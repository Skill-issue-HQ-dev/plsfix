// Direct unit coverage of summarize's own line, which the integration suites
// only exercise indirectly through a real updateLinks summary.
import { describe, expect, it } from "vitest";
import { summarize } from "./links";

describe("summarize", () => {
  it("counts links this computer holds no copy of", () => {
    expect(
      summarize({
        updated: 1,
        current: 0,
        notPasted: 2,
        inDeck: 0,
        missing: 0,
        wrongKey: 0,
        failed: 0,
        sourceChanges: [],
        failures: [],
        notes: [],
      }),
    ).toBe("1 updated, 2 not pasted yet");
  });

  it("says the rows the deck answered are up to date with this deck", () => {
    expect(
      summarize({
        updated: 0,
        current: 1,
        notPasted: 0,
        inDeck: 3,
        missing: 0,
        wrongKey: 0,
        failed: 0,
        sourceChanges: [],
        failures: [],
        notes: [],
      }),
    ).toBe("1 up to date, 3 up to date with this deck");
  });
});
