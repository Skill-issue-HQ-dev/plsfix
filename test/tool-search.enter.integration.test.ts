// @vitest-environment jsdom
// Find a tool in the BOOTED pane (the shipped taskpane.html and src/main.ts over
// the strict fake Excel host, relay mocked). The usability pass typed "multiple"
// and pressed Enter, and Enter pressed the first row, "Insert Trading comps",
// which writes a template block into the sheet. Enter now goes to a control and
// writes nothing; a second Enter, on the focused control, presses it the normal
// way, through the real guard and dispatch table.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
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
// pane calls the promise style office.js also supports.
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

beforeEach(async () => {
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
});

afterEach(() => {
  uninstallFakeHost();
});

const box = (): HTMLInputElement =>
  document.getElementById("tool-search") as HTMLInputElement;
const toast = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const rows = (): (string | undefined)[] =>
  [...document.querySelectorAll<HTMLElement>("#tool-search-results li")].map(
    (row) => row.dataset.toolAction,
  );
const snapshot = (): string =>
  JSON.stringify([helpers.cellMap("Model"), helpers.cellMap("Data")]);

function type(value: string): void {
  box().value = value;
  box().dispatchEvent(new Event("input", { bubbles: true }));
}

function pressInBox(key: string): void {
  box().dispatchEvent(
    new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }),
  );
}

describe("Enter in Find a tool, in the booted pane", () => {
  it('on "multiple" writes nothing and leaves focus on a control; Enter again runs it', async () => {
    helpers.seed("Model!B2", [[1234.5]]);
    helpers.select("Model!B2");
    const before = snapshot();

    type("multiple");
    expect(rows()[0]).toBe("cycle-number-multiple");
    pressInBox("Enter");
    await settle();

    // Nothing reached the workbook: no block, no number format, no toast.
    expect(snapshot()).toBe(before);
    expect(helpers.numberFormat("Model!B2")).not.toContain("x");
    expect(toast()).toBe("");
    // The pane went to the 0.0x cycle and left focus on it.
    const cycle = document.querySelector<HTMLButtonElement>(
      '[data-action="cycle-number-multiple"]',
    )!;
    expect(document.activeElement).toBe(cycle);
    expect(cycle.classList.contains("tool-found")).toBe(true);
    expect(box().value).toBe("");
    expect(document.getElementById("tool-search-results")!.hidden).toBe(true);

    // Enter again, on the focused control: pressed the normal way.
    pressEnterOn(document.activeElement as HTMLElement);
    await settle();
    expect(toast()).toBe("Number format: 0.0x");
    expect(helpers.numberFormat("Model!B2")).toBe('0.0"x";[Red](0.0"x");-');
  });

  it("takes you to a control on another tab and writes nothing there either", async () => {
    const before = snapshot();
    type("unhide");
    pressInBox("Enter");
    await settle();

    expect(snapshot()).toBe(before);
    expect(
      document.getElementById("tab-workbook")!.getAttribute("aria-selected"),
    ).toBe("true");
    expect(document.activeElement).toBe(
      document.querySelector('[data-action="sheets-unhide-all"]'),
    );
    expect(toast()).toBe("");
  });

  it("still lets the arrow keys choose which result Enter goes to", async () => {
    type("send to powerpoint");
    expect(rows().slice(0, 2)).toEqual(["export-selection", "export-chart"]);
    pressInBox("ArrowDown");
    pressInBox("Enter");
    await settle();

    expect(document.activeElement).toBe(
      document.getElementById("export-chart"),
    );
    expect(
      document.getElementById("tab-links")!.getAttribute("aria-selected"),
    ).toBe("true");
  });

  it("a click on a result goes to it and writes nothing", async () => {
    helpers.select("Model!A1");
    const before = snapshot();
    type("Insert WACC");
    document
      .querySelector<HTMLElement>("#tool-search-results li")!
      .dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await settle();

    expect(snapshot()).toBe(before);
    expect(document.activeElement).toBe(
      document.querySelector('[data-action="template-wacc"]'),
    );
  });

  it("says so, in one line, when nothing matches, and writes nothing", async () => {
    const before = snapshot();
    type("frobnicate");
    pressInBox("Enter");
    await settle();

    expect(snapshot()).toBe(before);
    expect(
      document.getElementById("tool-search-empty")!.textContent?.trim(),
    ).toBe('No tool matches "frobnicate". Try export, undo, color or dcf.');
  });
});
