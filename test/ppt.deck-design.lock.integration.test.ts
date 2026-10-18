// The design lock: every flow that writes the deck's design part (a palette
// commit, Put in the deck theme, Restore) runs inside it, so a commit made while
// a put is running can never write the part back without the colors the put just
// saved. One at a time, in the order asked, a failure never blocks the next.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { installFakePpt, uninstallFakePpt } from "./fakeppt";
import type * as DesignModule from "../src/ppt/deck-design";

let design: typeof DesignModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  installFakePpt({ slides: 1 });
  design = await import("../src/ppt/deck-design");
});
afterEach(() => {
  uninstallFakePpt();
});

describe("withDesignLock", () => {
  it("runs one design write after another, never two at once", async () => {
    const order: string[] = [];
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    const first = design.withDesignLock(async () => {
      order.push("first begins");
      await gate;
      order.push("first ends");
      return 1;
    });
    const second = design.withDesignLock(() => {
      order.push("second begins");
      return Promise.resolve(2);
    });
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(order).toEqual(["first begins"]);
    release();
    expect(await Promise.all([first, second])).toEqual([1, 2]);
    expect(order).toEqual(["first begins", "first ends", "second begins"]);
  });

  it("lets the next one run after one that failed, and still reports the failure", async () => {
    const failed = design.withDesignLock(() =>
      Promise.reject(new Error("the first one failed")),
    );
    const next = design.withDesignLock(() => Promise.resolve("ran"));
    await expect(failed).rejects.toThrow("the first one failed");
    expect(await next).toBe("ran");
  });

  it("hands each caller its own answer", async () => {
    const [a, b, c] = await Promise.all([
      design.withDesignLock(() => Promise.resolve("a")),
      design.withDesignLock(() => Promise.resolve("b")),
      design.withDesignLock(() => Promise.resolve("c")),
    ]);
    expect([a, b, c]).toEqual(["a", "b", "c"]);
  });
});
