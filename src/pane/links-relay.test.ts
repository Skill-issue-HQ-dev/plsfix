// What the Links tab may ask of the relay, by transport. The adapter is mocked
// (the flows themselves are proven over the real adapters in
// test/links.relay-inbox.integration.test.ts and
// test/links.local-no-relay.integration.test.ts); this file pins the two
// decisions: whether a relay push announces, and which relay Remove link uses.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { pushLinks } from "../excel";
import { LocalCollector } from "../link/local-collector";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { pushThroughRelay, removalRelay } from "./links-relay";

vi.mock("../excel", () => ({ pushLinks: vi.fn() }));

const SUMMARY = { pushed: 2, missing: 0, failed: 0, failures: [] };
const relay = {
  deleteLink: vi.fn(),
} as unknown as RelayApi;
const ws = { id: "ws" } as unknown as Workspace;

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(pushLinks).mockResolvedValue(SUMMARY);
});

describe("pushThroughRelay", () => {
  it("asks for the Inbox rows when this computer holds the link key", async () => {
    expect(await pushThroughRelay(["a", "b"], relay, ws)).toBe(SUMMARY);
    expect(pushLinks).toHaveBeenCalledWith(["a", "b"], relay, { announce: ws });
  });

  it("pushes revisions only when there is no key", async () => {
    await pushThroughRelay("all", relay, null);
    expect(pushLinks).toHaveBeenCalledWith("all", relay);
  });
});

describe("removalRelay", () => {
  it("is the real relay in relay mode, so its revoke still runs", () => {
    expect(removalRelay("relay", relay)).toBe(relay);
  });

  it("is a collector in local mode: the revoke does nothing and the relay is never touched", async () => {
    const standIn = removalRelay("local", relay);

    expect(standIn).toBeInstanceOf(LocalCollector);
    await expect(standIn.deleteLink("id", "auth")).resolves.toBeUndefined();
    expect(relay.deleteLink).not.toHaveBeenCalled();
  });
});
