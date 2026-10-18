// @vitest-environment jsdom
// The public build (no relay, src/link/local-only.ts) on the shipped pptpane.html over the fake
// host: a machine that chose the relay and holds a link key still boots into copy and paste, the
// Link key section is hidden, and the pane neither calls a relay nor fetches.

import { afterEach, describe, expect, it, vi } from "vitest";
import { bootInMode, modeRelay, relayCalls } from "./ppt.mode.support";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return modeRelay();
    },
  };
});

afterEach(() => {
  uninstallFakePpt();
  vi.unstubAllGlobals();
});

describe("the public build's PowerPoint pane", () => {
  it("boots into copy and paste on a machine that chose the relay, with no way to it", async () => {
    vi.stubGlobal("__PLSFIX_LOCAL_ONLY__", true);
    const fetches: string[] = [];
    vi.stubGlobal("fetch", (input: unknown) => {
      fetches.push(String(input));
      return Promise.reject(new Error("no network in the public build"));
    });

    await bootInMode("relay");

    const select = document.getElementById("link-transport");
    expect(select?.closest("section")?.hidden).toBe(true);
    expect((select as HTMLSelectElement).value).toBe("local");
    expect(document.getElementById("paste-links")?.hidden).toBe(false);
    expect(relayCalls()).toEqual([]);
    expect(fetches).toEqual([]);
  });
});
