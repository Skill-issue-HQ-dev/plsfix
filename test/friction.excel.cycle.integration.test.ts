// @vitest-environment jsdom
// The friction gate for the style-cycle key: the real Excel pane over the strict fake host, the
// handler the add-in registered for Ctrl+Shift+Alt+1 run as one action against
// test/friction-targets.ts, the HOST's state read back (each selected cell's style and the fill it
// resolves to) and the receipt the pane's toast shows.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import { actions, begin, finish, installScoreboard } from "./friction.support";
import { freshWebStorage } from "./web-storage";

enableStrictLoadSemantics();
installScoreboard();

const PRIMARY = "#111111";
const SHORTCUT = "Ctrl+Shift+Alt+1";
const COMMAND = "PLSFIX_CYC_STYLE";
const SELECTED = [
  "Model!B2",
  "Model!C2",
  "Model!D2",
  "Model!B3",
  "Model!C3",
  "Model!D3",
];

let helpers: FakeHelpers;

// A machine's web storage that works on every Node the suite runs on (Node 25 and up ship a
// localStorage global with no working methods).
function stubStorage(): void {
  const store = new Map<string, string>();
  vi.stubGlobal("localStorage", {
    getItem: (name: string) => store.get(name) ?? null,
    setItem: (name: string, value: string) => store.set(name, value),
    removeItem: (name: string) => store.delete(name),
    clear: () => store.clear(),
  });
}

// The fake host's Office.onReady only knows the callback style; src/main.ts calls the promise
// style too (the same patch the click-through suite makes).
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

// The latch is on while every tab button is disabled: wait it out, bounded.
async function settle(): Promise<void> {
  await drain();
  const busy = (): boolean => {
    const tabs = document.querySelectorAll<HTMLButtonElement>("[role=tab]");
    return tabs.length > 0 && [...tabs].every((tab) => tab.disabled);
  };
  for (let i = 0; i < 200 && busy(); i += 1) await drain(1);
}

async function bootExcel(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  stubStorage();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
  helpers.seed("Model!A1", [
    ["Revenue", 100, 110, 121],
    ["Cost", 40, 44, 48],
    ["Margin", 60, 66, 73],
  ]);
  helpers.select("Model!B2:D3");
}

const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const styleOf = (address: string): string => helpers.cell(address).style;

// What Office does on the key: it runs the handler the add-in registered for the command.
function fire(): ReturnType<typeof vi.fn> {
  const handler = helpers.actions().get(COMMAND);
  if (handler === undefined) throw new Error(`no handler for ${COMMAND}`);
  const completed = vi.fn();
  handler({ completed });
  return completed;
}

async function press(): Promise<void> {
  const completed = fire();
  await settle();
  expect(completed).toHaveBeenCalledOnce();
}

async function addStarters(): Promise<void> {
  document
    .querySelector<HTMLButtonElement>('[data-action="styles-starter"]')!
    .click();
  await settle();
}

beforeEach(async () => {
  await bootExcel();
});

afterEach(() => {
  uninstallFakeHost();
  vi.unstubAllGlobals();
});

describe("the style-cycle key", () => {
  it("steps the selection through the styles with one key: the first, on the whole selection", async () => {
    // Set up, not counted: the eight starters exist and the selection wears Normal.
    await addStarters();
    helpers.select("Model!B2:D3");

    begin("excel-style-cycle");
    actions.shortcut(SHORTCUT, fire);
    await settle();

    for (const address of SELECTED) {
      expect(styleOf(address), address).toBe("pls,fix Title");
      expect(helpers.fill(address).color, address).toBe(PRIMARY);
    }
    expect(styleOf("Model!A1")).toBe("Normal");
    expect(toastText()).toBe("Style: pls,fix Title (1 of 8)");
    finish("excel-style-cycle");
  });

  it("goes on a step at a time, through Normal, and round to the first again", async () => {
    await addStarters();
    helpers.select("Model!B2:D3");

    const seen: { style: string; toast: string }[] = [];
    for (let step = 0; step < 10; step += 1) {
      await press();
      seen.push({ style: styleOf("Model!B2"), toast: toastText() });
    }

    expect(seen.map((entry) => entry.style)).toEqual([
      "pls,fix Title",
      "pls,fix Header",
      "pls,fix Input",
      "pls,fix Formula",
      "pls,fix Result",
      "pls,fix Total",
      "pls,fix Note",
      "pls,fix Highlight",
      "Normal",
      "pls,fix Title",
    ]);
    expect(seen[1]?.toast).toBe("Style: pls,fix Header (2 of 8)");
    expect(seen[8]?.toast).toBe("Style: Normal");
    for (const address of SELECTED)
      expect(styleOf(address)).toBe("pls,fix Title");
  });

  it("with no style saved, says what to do in the pane's toast and changes nothing", async () => {
    await press();

    expect(toastText()).toBe("Save a style first: Tools > Custom styles.");
    expect(styleOf("Model!B2")).toBe("Normal");
  });
});
