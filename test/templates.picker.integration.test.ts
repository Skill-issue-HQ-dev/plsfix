// @vitest-environment jsdom
// The Templates picker in the BOOTED pane: the shipped taskpane.html and
// src/main.ts over the strict fake Excel host, relay mocked. What the picker's
// own unit test cannot show: a press runs through the real guard and dispatch
// table and writes the block, Enter in the search does the same, the "?" card
// lists every template, and the "Find a tool" box finds a template by name and
// runs it. Nothing here re-tests what a block computes (its own file does).

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
import { HELP } from "../src/help/copy";
import { TEMPLATES } from "../src/templates/index";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import { FakeRelay } from "./fakerelay";
import { pressEnterOn } from "./native-enter";
import { freshWebStorage } from "./web-storage";

enableStrictLoadSemantics();

let relay: FakeRelay;
vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

let helpers: FakeHelpers;

// test/fakehost.ts's Office.onReady only implements the callback style; the
// pane calls the promise style office.js also supports (same patch as
// test/clickthrough.excel.integration.test.ts).
function patchOnReadyForPromiseStyle(): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
}

async function drain(rounds = 12): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

// Busy shows as every tab disabled (the blanket latch); wait it out, bounded.
async function settle(): Promise<void> {
  await drain();
  for (let i = 0; i < 200; i += 1) {
    const tabs = document.querySelectorAll<HTMLButtonElement>("[role=tab]");
    if (!(tabs.length > 0 && [...tabs].every((tab) => tab.disabled))) return;
    await drain(1);
  }
}

async function bootPane(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  relay = new FakeRelay();
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
}

const toast = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const search = (): HTMLInputElement =>
  document.getElementById("template-search") as HTMLInputElement;

function type(box: HTMLInputElement, value: string): void {
  box.value = value;
  box.dispatchEvent(new Event("input", { bubbles: true }));
}

function pressKey(box: HTMLElement, key: string): void {
  box.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
}

beforeEach(async () => {
  await bootPane();
});

afterEach(() => {
  uninstallFakeHost();
});

describe("the Templates picker in the booted pane", () => {
  it("lists one Insert button per template, in the section the help and the rig key on", () => {
    const section = document
      .getElementById("templates-heading")
      ?.closest("section");
    const actions = [
      ...(section?.querySelectorAll<HTMLButtonElement>("button[data-action]") ??
        []),
    ].map((button) => button.dataset.action);
    expect(actions).toEqual(TEMPLATES.map((entry) => entry.id));
  });

  it("writes the block at the active cell when its Insert button is pressed", async () => {
    helpers.select("Model!C5");
    document
      .querySelector<HTMLButtonElement>('[data-action="template-wacc"]')
      ?.click();
    await settle();

    expect(toast()).toBe("Template written: WACC (21x2)");
    expect(helpers.value("Model!C5")).toBe("WACC");
    expect(helpers.formula("Model!D9")).toBe("=D7*(1+(1-D15)*D8)");
    expect(helpers.formula("Model!D12")).toBe("=D6+D9*D10+D11");
  });

  it("writes every template the picker lists, each over a clear sheet", async () => {
    for (const entry of TEMPLATES) {
      await bootPane();
      helpers.select("Model!B3");
      document
        .querySelector<HTMLButtonElement>(`[data-action="${entry.id}"]`)
        ?.click();
      await settle();
      expect(toast(), entry.id).toBe(
        `Template written: ${entry.name} (${String(entry.block.rows)}x${String(entry.block.cols)})`,
      );
    }
  });

  it("refuses a second block over the first, by name, and keeps the first", async () => {
    helpers.select("Model!C5");
    document
      .querySelector<HTMLButtonElement>('[data-action="template-wacc"]')
      ?.click();
    await settle();
    helpers.select("Model!C5");
    document
      .querySelector<HTMLButtonElement>('[data-action="template-precedents"]')
      ?.click();
    await settle();

    expect(toast()).toMatch(/Templates need an empty block of 19 rows by 6/);
    expect(helpers.value("Model!C5")).toBe("WACC");
  });

  it("inserts the first visible match when Enter is pressed in the search", async () => {
    helpers.select("Model!A1");
    type(search(), "football");
    pressKey(search(), "Enter");
    await settle();

    expect(toast()).toBe("Template written: Football field data (6x4)");
    expect(helpers.value("Model!A1")).toBe("Method (EUR m)");
  });

  it("says nothing and writes nothing on Enter when no template matches", async () => {
    helpers.select("Model!A1");
    type(search(), "zzz");
    pressKey(search(), "Enter");
    await settle();

    expect(helpers.value("Model!A1")).toBe("");
    expect(document.getElementById("template-empty")?.hidden).toBe(false);
  });

  it("lists every template, with its sentence, in the section's ? card", () => {
    document
      .getElementById("help-templates-heading-toggle")
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    const card = document.getElementById("help-templates-heading");
    expect(card?.hidden).toBe(false);

    const terms = [...(card?.querySelectorAll("dt") ?? [])].map(
      (node) => node.textContent,
    );
    expect(terms).toEqual(TEMPLATES.map((entry) => `Insert ${entry.name}`));
    const sentences = [...(card?.querySelectorAll("dd") ?? [])].map(
      (node) => node.textContent,
    );
    expect(sentences).toEqual(
      TEMPLATES.map(
        (entry) => HELP["templates-heading"]?.buttons[entry.id] ?? "missing",
      ),
    );
  });

  // Find a tool takes you to the Insert button and stops: a second Enter, now
  // on the focused button, is what writes the block.
  it("finds a template from the Find a tool box, goes to it on Enter and inserts it on a second Enter", async () => {
    helpers.select("Model!A1");
    const box = document.getElementById("tool-search") as HTMLInputElement;
    type(box, "Insert Trading comps");
    const first = document.querySelector<HTMLElement>(
      "#tool-search-results li",
    );
    expect(first?.dataset.toolAction).toBe("template-trading-comps");

    pressKey(box, "Enter");
    await settle();
    expect(helpers.value("Model!A1")).toBe("");
    expect(toast()).not.toMatch(/Template written/);
    expect(document.activeElement).toBe(
      document.querySelector('[data-action="template-trading-comps"]'),
    );

    pressEnterOn(document.activeElement as HTMLElement);
    await settle();
    expect(toast()).toBe("Template written: Trading comps (25x9)");
    expect(helpers.value("Model!A1")).toBe("Trading comps");
  });

  // The picker's own filter hides the WACC row; Find a tool must still reach
  // it, which takes that filter's words away so the row is there to land on.
  it("finds a template the picker's own search is hiding, goes to it and inserts it on a second Enter", async () => {
    helpers.select("Model!A1");
    type(search(), "comps");
    const row = document.querySelector<HTMLElement>(
      '[data-template="template-wacc"]',
    );
    expect(row?.hidden).toBe(true);

    const box = document.getElementById("tool-search") as HTMLInputElement;
    type(box, "Insert WACC");
    const first = document.querySelector<HTMLElement>(
      "#tool-search-results li",
    );
    expect(first?.dataset.toolAction).toBe("template-wacc");

    pressKey(box, "Enter");
    await settle();
    expect(helpers.value("Model!A1")).toBe("");
    expect(row?.hidden).toBe(false);
    expect(search().value).toBe("");
    expect(document.activeElement).toBe(
      document.querySelector('[data-action="template-wacc"]'),
    );

    pressEnterOn(document.activeElement as HTMLElement);
    await settle();
    expect(toast()).toBe("Template written: WACC (21x2)");
    expect(helpers.value("Model!A1")).toBe("WACC");
  });

  it("writes nothing on Enter in an empty search", async () => {
    helpers.select("Model!A1");
    pressKey(search(), "Enter");
    await settle();

    expect(helpers.value("Model!A1")).toBe("");
    expect(toast()).not.toMatch(/Template written/);
  });
});
