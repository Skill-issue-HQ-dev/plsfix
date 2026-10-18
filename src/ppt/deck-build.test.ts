// What a failed flow says about the record that rode with it: a timeout and a
// missing shape are read through the sentences that wrap them, anything else is
// the host refusing.

import { describe, expect, it } from "vitest";
import { ChartDrawTimeout } from "./chart-draw";
import { outcomeOf } from "./deck-build";
import { missingShapeError } from "./missing-shape";

const notFound = Object.assign(new Error("ItemNotFound: no part"), {
  code: "ItemNotFound",
});

describe("outcomeOf", () => {
  it("calls a host that stopped answering a timeout", () => {
    expect(outcomeOf(new ChartDrawTimeout("refreshing the text"))).toBe(
      "timeout",
    );
  });

  it("reads a timeout through the sentence that wraps it", () => {
    const wrapped = new Error("refresh Model!B4:F12 text: stopped", {
      cause: new ChartDrawTimeout("refreshing the text"),
    });
    expect(outcomeOf(wrapped)).toBe("timeout");
  });

  it("calls an ItemNotFound missing, before and after it is reworded", () => {
    expect(outcomeOf(notFound)).toBe("missing");
    expect(outcomeOf(missingShapeError("break x", notFound))).toBe("missing");
  });

  it("calls any other refusal failed", () => {
    expect(outcomeOf(new Error("GeneralException: the host refused"))).toBe(
      "failed",
    );
    expect(outcomeOf("a plain string")).toBe("failed");
  });

  it("stops reading causes at a sensible depth", () => {
    let chain: Error = new ChartDrawTimeout("x");
    for (let at = 0; at < 8; at += 1) chain = new Error("w", { cause: chain });
    expect(outcomeOf(chain)).toBe("failed");
  });
});
