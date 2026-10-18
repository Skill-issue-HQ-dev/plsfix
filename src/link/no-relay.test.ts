// src/link/no-relay.test.ts
// The local-only relay: every call refused in its own words and no request made; relayFor()
// hands it out only in a local-only build.
import { afterEach, describe, expect, it, vi } from "vitest";
import { NO_RELAY, NoRelay, relayFor } from "./no-relay";
import { RelayClient } from "./relay";

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("NoRelay", () => {
  it("refuses every call in words and never fetches", async () => {
    const fetchSpy = vi.fn();
    vi.stubGlobal("fetch", fetchSpy);
    const relay = new NoRelay();
    const calls = [
      relay.putLink("id", "auth", new Uint8Array(1)),
      relay.getLink("id", "auth"),
      relay.getLinkRev("id", "auth", 1),
      relay.deleteLink("id", "auth"),
      relay.status([]),
      relay.touchLinks([]),
      relay.fetchLinks([]),
      relay.postInbox("ws", "auth", "id", new Uint8Array(1)),
      relay.listInbox("ws", "auth"),
      relay.deleteInbox("ws", "auth", "id"),
    ];
    for (const call of calls) await expect(call).rejects.toThrow(NO_RELAY);
    expect(fetchSpy).not.toHaveBeenCalled();
  });
});

describe("relayFor", () => {
  it("gives the real client in the hosted build", () => {
    expect(relayFor("https://x.test/modelis/taskpane.html")).toBeInstanceOf(
      RelayClient,
    );
  });

  it("gives NoRelay in a local-only build", () => {
    vi.stubGlobal("__PLSFIX_LOCAL_ONLY__", true);
    expect(relayFor("https://x.test/plsfix/taskpane.html")).toBeInstanceOf(
      NoRelay,
    );
  });
});
