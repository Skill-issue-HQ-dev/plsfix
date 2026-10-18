// @vitest-environment jsdom
// The key hints in the BOOTED Excel pane: the shipped taskpane.html and
// src/main.ts over the strict fake Excel host, relay mocked. On Excel for the
// web (the fake's `web` option names itself OfficeOnline) the two buttons whose
// default keys the browser keeps name their web keys; on any other host the
// markup stays as written. The rewrite itself is src/pane/platform-keys.test.ts.

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

async function bootPane(web: boolean): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  installFakeHost({ sheets: ["Model"], web });
  relay = new FakeRelay();
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
}

const titleOf = (action: string): string | null =>
  document.querySelector(`[data-action="${action}"]`)?.getAttribute("title") ??
  null;

afterEach(() => {
  uninstallFakeHost();
});

describe("the key hints in the booted pane", () => {
  it("name the web keys on Excel for the web", async () => {
    await bootPane(true);

    expect(titleOf("paint-apply-2")).toBe(
      "Paint slot 2 over the selection (Ctrl+Shift+Alt+2)",
    );
    expect(titleOf("cycle-col-width")).toBe(
      "Column widths: 64 / 80 / 96 / 120 / 48 pt (Ctrl+Shift+Alt+W)",
    );
    // A button whose key the browser leaves alone keeps it.
    expect(titleOf("paint-apply-1")).toBe(
      "Paint slot 1 over the selection (Ctrl+Shift+D)",
    );
  });

  it("stay as written on any other host", async () => {
    await bootPane(false);

    expect(titleOf("paint-apply-2")).toBe(
      "Paint slot 2 over the selection (Ctrl+Shift+N)",
    );
    expect(titleOf("cycle-col-width")).toBe(
      "Column widths: 64 / 80 / 96 / 120 / 48 pt (Ctrl+Shift+X)",
    );
  });
});
