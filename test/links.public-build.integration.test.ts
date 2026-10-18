// @vitest-environment jsdom
// The public build (no relay, src/link/local-only.ts) on the real Links tab over the real
// adapters: a device that chose the relay and holds a link key still boots into copy and paste,
// the Transport and Link key sections are hidden, and an export plus Copy all reach no relay and
// make no fetch.

import { afterEach, describe, expect, it, vi } from "vitest";
import { closeWorkbook, openWorkbook } from "./links-tab.support";

afterEach(() => {
  closeWorkbook();
  vi.unstubAllGlobals();
});

describe("the public build's Links tab", () => {
  it("stays local on a device that chose the relay, and offers no way to it", async () => {
    vi.stubGlobal("__PLSFIX_LOCAL_ONLY__", true);
    const rig = await openWorkbook({ transport: "relay", key: true });
    await rig.mount();

    const transport = document
      .getElementById("link-transport")
      ?.closest("section");
    expect(transport?.hidden).toBe(true);
    expect(document.getElementById("link-key-section")?.hidden).toBe(true);
    expect(document.getElementById("push-selected")?.textContent).toContain(
      "Copy selected",
    );

    rig.helpers.seed("Model!B8", [
      [1, 2],
      [3, 4],
    ]);
    rig.helpers.select("Model!B8:C9");
    rig.click("export-selection");
    await rig.settle();
    rig.click("push-all");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.calls).toEqual([]);
    expect(rig.fetches).toEqual([]);
  });
});
