// @vitest-environment jsdom
// L3 through the real pane: the shipped pptpane.html over the fake deck and the
// fake relay. A link the relay forgot (past its 30 days) reads In this deck with
// the time its data was sent, Update all paints nothing and says so, and Revert
// says in words why it cannot go back. Strict load semantics.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { insertFirstWaiting, waitingText } from "./ppt.deck-rows.support";
import {
  bootPane,
  click,
  copiedDetails,
  deck,
  linkRows,
  paneRelay,
  settle,
  tickRow,
  toastText,
} from "./stress.ppt.support";
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

beforeEach(() => {
  vi.useFakeTimers({ toFake: ["Date", "performance"] });
  vi.setSystemTime(new Date(2026, 9, 2, 14, 2, 0));
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a link the relay forgot, in the pane", () => {
  it("reads In this deck with its data time, and Update all and Revert answer in words", async () => {
    await bootPane();
    const item = await waitingText("EUR 3.2m");
    await insertFirstWaiting();
    expect(linkRows()[0]!.querySelector(".badge")!.textContent).toBe(
      "Up to date",
    );
    paneRelay().links.delete(item.id);

    click("refresh-links");
    await settle();
    const row = linkRows()[0]!;
    expect(row.querySelector(".badge")!.textContent).toBe("In this deck");
    expect(row.querySelector(".link-status .link-meta")!.textContent).toBe(
      "data from 02.10 14:02",
    );
    expect(row.querySelector(".link-updated")!.textContent).toBe("just now");

    click("update-all");
    await settle();
    expect(toastText()).toBe("1 up to date with this deck");
    expect(deck().slides[0]!.shapes[0]!.text).toBe("EUR 3.2m");

    tickRow(0);
    click("revert-selected");
    await settle();
    expect(toastText()).toBe("1 without a previous version");
    expect(await copiedDetails()).toContain(
      "this deck keeps only the version on the slide",
    );
    expect(deck().slides[0]!.shapes[0]!.text).toBe("EUR 3.2m");
  });
});
