// @vitest-environment jsdom
// Apply and Reset all in the BOOTED Excel pane while Office has not loaded
// pls,fix's key table: the shipped taskpane.html and src/main.ts over the strict
// fake Excel host, relay mocked. The busy latch hands every button back when an
// action ends, so main.ts's rule-button sync has to put these two back off; with
// the table loaded they come back on like any other button.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type * as RelayModule from "../src/link/relay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import { FakeRelay } from "./fakerelay";
import {
  loadedKeys,
  sectionButtons,
  serveShortcuts,
} from "./shortcuts.support";
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

const NOT_LOADED =
  "Office has not loaded pls,fix's keyboard shortcuts here, so the keys do nothing in this window. The buttons and the ribbon work as usual.";

// Same patch as test/templates.picker.integration.test.ts: the fake's
// Office.onReady only implements the callback style, the pane uses the promise.
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

// The key table Office answers with: what it holds for the add-in, or nothing.
async function bootPane(table: Record<string, string | null>): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  installFakeHost({ sheets: ["Model"] });
  const office = (
    globalThis as unknown as { Office: { actions: Record<string, unknown> } }
  ).Office;
  Object.assign(office.actions, {
    getShortcuts: async () => table,
    replaceShortcuts: async () => undefined,
    areShortcutsInUse: async () => [],
  });
  serveShortcuts();
  relay = new FakeRelay();
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
}

const note = (): string =>
  (document.getElementById("shortcuts-note")?.textContent ?? "").trim();

// Any action that goes through the guard: Undo with nothing to undo.
async function pressUndo(): Promise<HTMLButtonElement> {
  const undo = document.querySelector<HTMLButtonElement>(
    '[data-action="undo"]',
  )!;
  undo.click();
  await settle();
  return undo;
}

afterEach(() => {
  uninstallFakeHost();
});

describe("Apply and Reset all in the booted pane", () => {
  it("stay off after an action ends while Office has not loaded the key table", async () => {
    await bootPane({});
    expect(note()).toBe(NOT_LOADED);
    expect(sectionButtons().apply.disabled).toBe(true);
    expect(sectionButtons().reset.disabled).toBe(true);

    const undo = await pressUndo();

    // The latch released: other buttons are back on, these two are not.
    expect(undo.disabled).toBe(false);
    expect(sectionButtons().apply.disabled).toBe(true);
    expect(sectionButtons().reset.disabled).toBe(true);
    expect(note()).toBe(NOT_LOADED);
  });

  it("come back on with every other button once the table is loaded", async () => {
    await bootPane(loadedKeys());
    expect(sectionButtons().apply.disabled).toBe(false);

    const undo = await pressUndo();

    expect(undo.disabled).toBe(false);
    expect(sectionButtons().apply.disabled).toBe(false);
    expect(sectionButtons().reset.disabled).toBe(false);
  });
});
