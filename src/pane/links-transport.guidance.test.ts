// @vitest-environment jsdom
// src/pane/links-transport.guidance.test.ts
// What the Excel pane tells a user in each link mode: how to update a deck, the
// privacy statement, how the Brand colors reach PowerPoint. The shipped
// taskpane.html, with installLinksTransport switching the mode.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
import { WORKSPACE_STORAGE_KEY, type KeyStore } from "../link/workspace";
import type { Toast } from "../ui/toast";
import { installLinksTransport, type LinksTransport } from "./links-transport";

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

const toast: Toast = { show: () => undefined };

// A fresh install boots local; a store that holds a link key boots relay.
async function boot(mode: "local" | "relay"): Promise<LinksTransport> {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const keyStore = memoryStore();
  if (mode === "relay")
    await keyStore.set(WORKSPACE_STORAGE_KEY, "k".repeat(43));
  return installLinksTransport({ root: document, keyStore, toast });
}

function line(id: string): HTMLElement {
  const found = document.getElementById(id);
  if (found === null) throw new Error(`no #${id} on the Excel pane`);
  return found;
}

function words(id: string): string {
  return (line(id).textContent ?? "").replace(/\s+/g, " ").trim();
}

beforeEach(() => {
  document.body.innerHTML = "";
});

describe("the Links tab's line about updating", () => {
  const UPDATE_LINE =
    "Changed the source? Press Copy all (or tick links and Copy selected), then paste in PowerPoint > Inbox. The paste updates the slides.";

  it("copy and paste: stays on screen under the Copy buttons", async () => {
    await boot("local");
    expect(line("links-update-hint").hidden).toBe(false);
    expect(words("links-update-hint")).toBe(UPDATE_LINE);
    const buttons = line("push-all").closest(".io-row");
    expect(buttons?.nextElementSibling).toBe(line("links-update-hint"));
  });

  it("relay: is not shown, because Push is how a relay user updates", async () => {
    await boot("relay");
    expect(line("links-update-hint").hidden).toBe(true);
  });

  it("follows the mode when it is switched, both ways", async () => {
    const transport = await boot("relay");
    await transport.setMode("local");
    expect(line("links-update-hint").hidden).toBe(false);
    await transport.setMode("relay");
    expect(line("links-update-hint").hidden).toBe(true);
  });
});

describe("the Links tab's privacy statement", () => {
  const RELAY =
    "Anyone holding a deck can pull a link's latest picture for up to 30 days. Break the links in PowerPoint before sending a deck outside.";
  const LOCAL =
    "Link data stays in this workbook, the deck and this computer. Nothing is kept on a server.";

  it("copy and paste: says nothing is kept on a server, not the 30 days", async () => {
    await boot("local");
    expect(line("links-privacy-local").hidden).toBe(false);
    expect(words("links-privacy-local")).toBe(LOCAL);
    expect(line("links-privacy-relay").hidden).toBe(true);
  });

  it("relay: keeps the 30-day statement, not the local one", async () => {
    await boot("relay");
    expect(line("links-privacy-relay").hidden).toBe(false);
    expect(words("links-privacy-relay")).toBe(RELAY);
    expect(line("links-privacy-local").hidden).toBe(true);
  });

  it("follows the mode when it is switched, both ways", async () => {
    const transport = await boot("relay");
    await transport.setMode("local");
    expect(line("links-privacy-relay").hidden).toBe(true);
    expect(line("links-privacy-local").hidden).toBe(false);
    await transport.setMode("relay");
    expect(line("links-privacy-relay").hidden).toBe(false);
    expect(line("links-privacy-local").hidden).toBe(true);
  });
});

describe("the Brand tab's note on how the colors reach PowerPoint", () => {
  const NOTE =
    "These colors go to PowerPoint with every copy from the Links tab (Copy all or Copy selected).";

  it("copy and paste: sits right under the brand colors", async () => {
    await boot("local");
    expect(line("brand-powerpoint-hint").hidden).toBe(false);
    expect(words("brand-powerpoint-hint")).toBe(NOTE);
    expect(line("read-selection-colors").nextElementSibling).toBe(
      line("brand-powerpoint-hint"),
    );
  });

  it("relay: is not shown, because no color travels through the relay", async () => {
    await boot("relay");
    expect(line("brand-powerpoint-hint").hidden).toBe(true);
  });

  it("follows the mode when it is switched, both ways", async () => {
    const transport = await boot("relay");
    await transport.setMode("local");
    expect(line("brand-powerpoint-hint").hidden).toBe(false);
    await transport.setMode("relay");
    expect(line("brand-powerpoint-hint").hidden).toBe(true);
  });
});
