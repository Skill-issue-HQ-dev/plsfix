// The deadline every PowerPoint round trip runs under, and which host it asks:
// PowerPoint for the web gives up on a silent sync at 35 s (the longest add
// sync that ever answered took 27.1 s, S1), every other host at 60 s.

import { afterEach, describe, expect, it, vi } from "vitest";
import { installFakePpt, uninstallFakePpt } from "../../test/fakeppt";
import { SYNC_TIMEOUT_MS } from "./chart-cost";
import {
  ChartDrawTimeout,
  drawHost,
  SYNC_TIMEOUT_MS as EXPORTED_TIMEOUTS,
  withSyncDeadline,
} from "./chart-draw";

afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

// Starts a round trip that never answers on a host of this platform and
// reports what the deadline has made of it after each step in `steps` ms.
async function silentSyncAfter(
  platform: string,
  steps: number[],
): Promise<string[]> {
  vi.useFakeTimers();
  const { helpers } = installFakePpt({ slides: 1 });
  helpers.setPlatform(platform);
  let outcome = "waiting";
  withSyncDeadline(new Promise<never>(() => undefined), "testing").then(
    () => (outcome = "answered"),
    (error: unknown) => {
      outcome = error instanceof ChartDrawTimeout ? "timed out" : "failed";
    },
  );
  const seen: string[] = [];
  for (const step of steps) {
    await vi.advanceTimersByTimeAsync(step);
    seen.push(outcome);
  }
  return seen;
}

describe("the sync deadline", () => {
  it("is 35 s on the web and 60 s everywhere else", () => {
    expect(SYNC_TIMEOUT_MS).toEqual({ web: 35_000, desktop: 60_000 });
    expect(EXPORTED_TIMEOUTS).toBe(SYNC_TIMEOUT_MS);
  });

  it("rejects a sync the web never answers at 35 s, not a moment before", async () => {
    const seen = await silentSyncAfter("OfficeOnline", [34_999, 1]);
    expect(seen).toEqual(["waiting", "timed out"]);
  });

  it("waits the full 60 s for a desktop PowerPoint", async () => {
    const seen = await silentSyncAfter("PC", [35_000, 24_999, 1]);
    expect(seen).toEqual(["waiting", "waiting", "timed out"]);
  });

  it("waits 60 s on a Mac too", async () => {
    const seen = await silentSyncAfter("Mac", [59_999, 1]);
    expect(seen).toEqual(["waiting", "timed out"]);
  });

  it("answers a sync that comes back in time with its own answer", async () => {
    vi.useFakeTimers();
    const { helpers } = installFakePpt({ slides: 1 });
    helpers.setPlatform("OfficeOnline");
    const work = withSyncDeadline(Promise.resolve("done"), "testing");
    await expect(work).resolves.toBe("done");
  });

  it("names the round trip that stopped", async () => {
    vi.useFakeTimers();
    const { helpers } = installFakePpt({ slides: 1 });
    helpers.setPlatform("OfficeOnline");
    const work = withSyncDeadline(
      new Promise<never>(() => undefined),
      "reading the slide",
    );
    const caught = work.catch((error: unknown) => error);
    await vi.advanceTimersByTimeAsync(35_000);
    await expect(caught).resolves.toMatchObject({
      message: "PowerPoint stopped answering while reading the slide",
    });
  });
});

describe("drawHost", () => {
  it("is the web on PowerPoint for the web, the desktop on every other platform", () => {
    const { helpers } = installFakePpt({ slides: 1 });
    helpers.setPlatform("OfficeOnline");
    expect(drawHost()).toBe("web");
    for (const platform of ["PC", "Mac", "Universal"]) {
      helpers.setPlatform(platform);
      expect(drawHost()).toBe("desktop");
    }
  });

  it("is the desktop where no Office host exists at all, so a bare unit test never throws", () => {
    uninstallFakePpt();
    expect(drawHost()).toBe("desktop");
  });
});
