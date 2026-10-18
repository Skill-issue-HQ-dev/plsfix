// @vitest-environment jsdom
// src/pane/links-transport.palette.test.ts
// An Excel copy carries the workbook's brand colors in its bundle, so a paste in
// PowerPoint can adopt them without retyping. Read here through the manual-copy
// box (the text a refused clipboard write falls back to).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { decodeBundle } from "../link/bundle";
import type { KeyStore } from "../link/workspace";
import { DEFAULT_SETTINGS, setActiveSettings } from "../settings";
import type { Toast } from "../ui/toast";
import { installLinksTransport } from "./links-transport";

const ID = "a".repeat(32);
const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";

class FakeItem {
  constructor(readonly items: Record<string, Promise<Blob>>) {}
}

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

// A copy whose clipboard write is refused, so the bundle is left in the manual
// box a person would copy by hand.
async function copiedBundle(): Promise<ReturnType<typeof decodeBundle>> {
  vi.stubGlobal("ClipboardItem", FakeItem);
  vi.stubGlobal("navigator", {
    clipboard: { write: async () => Promise.reject(new Error("refused")) },
  });
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const toast: Toast = { show: () => undefined };
  const transport = await installLinksTransport({
    root: document,
    keyStore: memoryStore(),
    toast,
  });
  await transport.copy(
    async (_ws, relay) => {
      await relay.putLink(ID, "auth", new Uint8Array([1]));
      await relay.postInbox("ws", "auth", ID, new Uint8Array([2]));
      return { label: "Model!A1" };
    },
    { copied: () => "Copied", ready: () => "Ready" },
  );
  expect(() => transport.retryCopy()).toThrow();
  const manual = document.getElementById("copy-manual") as HTMLTextAreaElement;
  return decodeBundle(manual.value);
}

afterEach(() => {
  setActiveSettings({ ...DEFAULT_SETTINGS });
  vi.unstubAllGlobals();
});

describe("what a copy carries of the brand", () => {
  it("is the workbook's brand list, in order", async () => {
    setActiveSettings({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: [GOLD, MARBLE, GREEN],
    });
    const read = await copiedBundle();
    expect(read.ok && read.bundle.palette).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("is each color once when primary and accent were set apart to the same color", async () => {
    setActiveSettings({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD, MARBLE, GREEN, MARBLE],
    });
    const read = await copiedBundle();
    expect(read.ok && read.bundle.palette).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("carries no palette at all when the workbook holds one color only", async () => {
    setActiveSettings({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD],
    });
    const read = await copiedBundle();
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.bundle).not.toHaveProperty("palette");
  });

  it("is the shipped two when the workbook has no brand of its own", async () => {
    setActiveSettings({ ...DEFAULT_SETTINGS });
    const read = await copiedBundle();
    expect(read.ok && read.bundle.palette).toEqual([
      DEFAULT_SETTINGS.primary,
      DEFAULT_SETTINGS.accent,
    ]);
  });

  it("is still a copy with its links in it", async () => {
    const read = await copiedBundle();
    expect(read.ok && read.bundle.links.map((link) => link.id)).toEqual([ID]);
  });
});
