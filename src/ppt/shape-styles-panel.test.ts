// @vitest-environment jsdom
// The Shape styles section's own wiring with no Office host and no pane: which of
// two reads in flight draws, and that rows born while the pane is busy start
// disabled. The presses over the real pane are in
// test/ppt.shape-styles.pane.integration.test.ts.

import { beforeEach, describe, expect, it, vi } from "vitest";
import type { KeyStore } from "../link/workspace";
import { DEFAULT_SETTINGS } from "../settings";
import { encodeShapeStyles } from "../shape-styles-codec";
import { installShapeStyles } from "./shape-styles-panel";

beforeEach(() => {
  document.body.innerHTML = `
    <button id="tab-tools"></button>
    <input id="shape-style-name" />
    <button id="save-shape-style"></button>
    <ul id="shape-style-list" hidden></ul>
    <p id="shape-style-empty"></p>`;
});

const rawList = (...names: string[]): string =>
  encodeShapeStyles(names.map((name) => ({ name, kind: "user" as const })));

// A store whose reads answer when the test says so, in any order.
function heldStore(): { store: KeyStore; answers: ((raw: string) => void)[] } {
  const answers: ((raw: string) => void)[] = [];
  const store: KeyStore = {
    get: () =>
      new Promise<string | null>((resolve) => {
        answers.push(resolve);
      }),
    set: () => Promise.resolve(),
    remove: () => Promise.resolve(),
  };
  return { store, answers };
}

const install = (store: KeyStore, busy: () => boolean = () => false) =>
  installShapeStyles({
    act: vi.fn(),
    busy,
    note: vi.fn(),
    keyStore: store,
    palette: () => DEFAULT_SETTINGS,
  });

const names = (): string[] =>
  [...document.querySelectorAll<HTMLElement>("#shape-style-list li")].map(
    (row) => row.dataset.style!,
  );
const buttons = (): HTMLButtonElement[] => [
  ...document.querySelectorAll<HTMLButtonElement>("#shape-style-list button"),
];

describe("two reads in flight", () => {
  it("draws only the newest one when the older answers last", async () => {
    const { store, answers } = heldStore();
    const panel = install(store);
    const older = panel.refresh();
    const newer = panel.refresh();
    answers[1]!(rawList("Newer"));
    await newer;
    expect(names()).toEqual(["Newer"]);
    answers[0]!(rawList("Older"));
    await older;
    expect(names()).toEqual(["Newer"]);
  });

  it("draws only the newest one when the older answers first", async () => {
    const { store, answers } = heldStore();
    const panel = install(store);
    const older = panel.refresh();
    const newer = panel.refresh();
    answers[0]!(rawList("Older"));
    await older;
    expect(names()).toEqual([]);
    answers[1]!(rawList("Newer"));
    await newer;
    expect(names()).toEqual(["Newer"]);
  });
});

describe("rows born while the pane is busy", () => {
  it("start disabled, and live when nothing runs", async () => {
    const { store, answers } = heldStore();
    let busy = true;
    const panel = install(store, () => busy);
    const first = panel.refresh();
    answers[0]!(rawList("Mine"));
    await first;
    expect(buttons()).toHaveLength(3);
    expect(buttons().every((button) => button.disabled)).toBe(true);
    busy = false;
    const second = panel.refresh();
    answers[1]!(rawList("Mine"));
    await second;
    expect(buttons().every((button) => !button.disabled)).toBe(true);
  });
});

describe("the Tools tab", () => {
  it("reads the list when it opens, except while an action runs", async () => {
    const { store, answers } = heldStore();
    let busy = true;
    install(store, () => busy);
    document.getElementById("tab-tools")!.click();
    expect(answers).toHaveLength(0);
    busy = false;
    document.getElementById("tab-tools")!.click();
    expect(answers).toHaveLength(1);
  });
});
