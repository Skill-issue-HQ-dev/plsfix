// @vitest-environment jsdom
// The shipped PowerPoint pane's boot order: the deck's link list is on screen
// before the brand colors are read, so a deck with a design part never makes the
// list wait one to three round trips for colors. The brand controller is replaced
// by a recorder; main.ts, the deck and the relay are the real ones (harness:
// stress.ppt.support).
import { afterEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { bootPane, linkRows, paneRelay, plant } from "./stress.ppt.support";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

afterEach(() => {
  vi.doUnmock("../src/ppt/brand-colors");
  uninstallFakePpt();
});

// What the pane's boot sees in place of the brand controller: it only notes when
// its boot step runs.
function recorder(onBoot: () => void): void {
  vi.doMock("../src/ppt/brand-colors", () => ({
    installBrandColors: () => ({
      boot: async () => {
        onBoot();
      },
      receive: async () => "",
      show: () => undefined,
    }),
    drawBrandColors: () => undefined,
  }));
}

describe("the pane's boot", () => {
  it("lists the deck's links first and reads the brand colors after them", async () => {
    const events: string[] = [];
    let observer: MutationObserver | undefined;
    await bootPane(true, async () => {
      await plant();
      recorder(() => events.push("brand colors read"));
      observer = new MutationObserver(() => {
        if (linkRows().length > 0 && !events.includes("link list drawn")) {
          events.push("link list drawn");
        }
      });
      observer.observe(document.getElementById("link-rows")!, {
        childList: true,
        subtree: true,
      });
    });
    observer?.disconnect();
    expect(linkRows()).toHaveLength(1);
    expect(events).toEqual(["link list drawn", "brand colors read"]);
  });

  it("still reads the brand colors once, whatever the deck holds", async () => {
    let reads = 0;
    await bootPane(true, () => {
      recorder(() => (reads += 1));
    });
    expect(reads).toBe(1);
  });
});
