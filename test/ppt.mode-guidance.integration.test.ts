// @vitest-environment jsdom
// The PowerPoint pane's words by link mode: copy and paste on a fresh install,
// the relay on a machine that holds a link key (New here card, empty Inbox,
// Update all, privacy line, brand reason). Drives the shipped pptpane.html.
// Invariant: relay mode keeps its words; local mode never names a link key.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  bootInMode,
  chooseMode,
  emptyInboxText,
  lineById,
  modeRelay,
  plantLink,
  press,
  relayCalls,
  shownText,
  toastText,
} from "./ppt.mode.support";
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

const LOCAL_CARD = [
  "New here?",
  "1. In Excel: Links > Export selection (it copies for PowerPoint).",
  "2. Here: click the box, press Ctrl+V, then Insert.",
  "3. Numbers changed? In Excel press Copy all and paste here again.",
];

const RELAY_CARD = [
  "New here?",
  "1. Paste the link key from Excel > Links > Link key once.",
  "2. Insert what Excel exported; Update all repaints it later.",
];

// The card's heading is an h2, not a paragraph: read both, in page order.
function cardLines(): string[] {
  return shownText("#first-run-ppt", "h2, p");
}

describe("a fresh install (copy and paste)", () => {
  it("the New here card walks through export, paste and update", async () => {
    await bootInMode("local");
    expect(cardLines()).toEqual(LOCAL_CARD);
  });

  it("an empty Inbox says to press Export in Excel and paste here", async () => {
    await bootInMode("local");
    expect(document.getElementById("inbox-list")!.hidden).toBe(false);
    expect(emptyInboxText()).toBe(
      "Nothing waiting. In Excel press Export, then paste here.",
    );
  });
});

describe("a machine that already holds a link key (relay)", () => {
  it("the New here card keeps its words about the link key", async () => {
    await bootInMode("relay");
    expect(cardLines()).toEqual(RELAY_CARD);
  });

  it("an empty Inbox keeps telling the user to push from Excel", async () => {
    await bootInMode("relay");
    expect(document.getElementById("inbox-list")!.hidden).toBe(false);
    expect(emptyInboxText()).toBe(
      "Nothing waiting. Push a range or a chart from Excel.",
    );
  });
});

describe("switching the mode in Settings", () => {
  it("changes the card and the empty Inbox line at once, both ways", async () => {
    await bootInMode("relay");

    await chooseMode("local");
    expect(cardLines()).toEqual(LOCAL_CARD);
    expect(emptyInboxText()).toBe(
      "Nothing waiting. In Excel press Export, then paste here.",
    );

    await chooseMode("relay");
    expect(cardLines()).toEqual(RELAY_CARD);
    expect(emptyInboxText()).toBe(
      "Nothing waiting. Push a range or a chart from Excel.",
    );
  });
});

const TAIL = "To update from Excel, copy there and paste on the Inbox tab.";

describe("Update all", () => {
  it("copy and paste: the receipt says how to update from Excel", async () => {
    await bootInMode("local");
    await press("update-all");
    expect(toastText()).toBe(`No links found. ${TAIL}`);
  });

  it("copy and paste: the same line follows the count of a real deck", async () => {
    const deck = await bootInMode("local");
    plantLink(deck);
    await press("refresh-links");
    await press("update-all");
    expect(toastText()).toMatch(/^1 [a-z ]+\. To update from Excel, /);
    expect(toastText().endsWith(TAIL)).toBe(true);
  });

  it("relay: the receipt stays as it was", async () => {
    await bootInMode("relay");
    await press("update-all");
    expect(toastText()).toBe("No links found");
  });
});

describe("the privacy line under Break link", () => {
  const RELAY =
    "Anyone holding this deck can pull a linked picture for 30 days. Break links before sending a deck outside.";
  const LOCAL =
    "Link data stays in this workbook, the deck and this computer. Nothing is kept on a server.";

  it("copy and paste: says nothing is kept on a server, not the 30 days", async () => {
    await bootInMode("local");
    expect(lineById("ppt-privacy-local")).toEqual({ shown: true, text: LOCAL });
    expect(lineById("ppt-privacy-relay").shown).toBe(false);
  });

  it("relay: keeps the 30-day line, not the local one", async () => {
    await bootInMode("relay");
    expect(lineById("ppt-privacy-relay")).toEqual({ shown: true, text: RELAY });
    expect(lineById("ppt-privacy-local").shown).toBe(false);
  });

  it("follows the mode in Settings, both ways", async () => {
    await bootInMode("relay");
    await chooseMode("local");
    expect(lineById("ppt-privacy-local").shown).toBe(true);
    expect(lineById("ppt-privacy-relay").shown).toBe(false);
    await chooseMode("relay");
    expect(lineById("ppt-privacy-local").shown).toBe(false);
    expect(lineById("ppt-privacy-relay").shown).toBe(true);
  });
});

describe("the reason under the four brand buttons", () => {
  const REASON =
    "Excel colors arrive with a copy: Excel > Links > Copy all, paste on the Inbox tab. Or paste hex codes above.";

  it("copy and paste: says how Excel's colors arrive, right under the buttons", async () => {
    await bootInMode("local");
    expect(lineById("ppt-brand-copy-hint")).toEqual({
      shown: true,
      text: REASON,
    });
    const buttons = document.getElementById(
      "use-workbook-colors",
    )!.parentElement;
    expect(buttons?.nextElementSibling).toBe(
      document.getElementById("ppt-brand-copy-hint"),
    );
  });

  it("relay: is not shown, because colors only travel with a copy", async () => {
    await bootInMode("relay");
    expect(lineById("ppt-brand-copy-hint").shown).toBe(false);
  });

  it("follows the mode in Settings, both ways", async () => {
    await bootInMode("relay");
    await chooseMode("local");
    expect(lineById("ppt-brand-copy-hint").shown).toBe(true);
    await chooseMode("relay");
    expect(lineById("ppt-brand-copy-hint").shown).toBe(false);
  });
});

describe("the local-mode statement is true of this pane", () => {
  it("the list, the Inbox and Update all reach neither the relay nor the network", async () => {
    const fetched: string[] = [];
    vi.stubGlobal("fetch", (url: unknown): Promise<never> => {
      fetched.push(String(url));
      return Promise.reject(new Error("no network in copy-and-paste mode"));
    });
    const deck = await bootInMode("local");
    plantLink(deck);
    await press("refresh-links");
    await press("refresh-inbox");
    await press("update-all");
    await press("paste-latest-linked");
    expect(relayCalls()).toEqual([]);
    expect(fetched).toEqual([]);
  });
});
