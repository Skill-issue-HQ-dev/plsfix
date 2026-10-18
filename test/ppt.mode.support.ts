// Harness for the ppt.mode-*.integration suites: the shipped pptpane.html in
// jsdom over the fake host, booted in a chosen link mode, plus the DOM reads
// that say what a user sees in that mode. Owns no assertions.
// Invariant: every boot is a fresh module graph, deck, relay and key store.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
import { newToken } from "../src/link/crypto";
import {
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
} from "../src/link/model";
import { TRANSPORT_STORAGE_KEY } from "../src/link/transport-setting";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import { settlePpt, trackPptBoot } from "./ppt-ready";
import { recordCalls } from "./relay-calls";

let relay: FakeRelay;
let calls: string[] = [];

// The pane builds its own RelayClient; each suite's vi.mock hands it this one.
export function modeRelay(): FakeRelay {
  return relay;
}

// Every method the pane has called on that relay since the boot.
export function relayCalls(): string[] {
  return calls;
}

const noStore: KeyStore = {
  get: async () => null,
  set: async () => undefined,
  remove: async () => undefined,
};

// A fresh pane. "local" is a new install (nothing stored, copy and paste);
// "relay" is a machine that already holds a link key. Answers the deck.
export async function bootInMode(
  mode: "local" | "relay",
): Promise<FakePresentation> {
  vi.resetModules();
  uninstallFakePpt();
  const recorded = recordCalls(new FakeRelay());
  relay = recorded.relay;
  calls = recorded.calls;
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>();
  if (mode === "relay") {
    const workspace = await createWorkspace(noStore);
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  }
  storage.set(TRANSPORT_STORAGE_KEY, mode);
  const host = installFakePpt({ slides: 3, storage });
  host.helpers.selectSlide(host.presentation.slides[0]!.id);
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
  await settlePpt();
  return host.presentation;
}

// One linked picture on slide 2, tagged the way an insert leaves it: the deck
// a colleague sent. The pane lists it once refresh-links reads the deck.
export function plantLink(deck: FakePresentation): void {
  const shape = deck.addShape(deck.slides[1]!, {
    type: "GeometricShape",
    fillImage: fakePng(800, 400),
    left: 100,
    top: 80,
    width: 400,
    height: 200,
  });
  const tag: LinkTag = {
    v: 1,
    id: newLinkId((n) => new Uint8Array(n).fill(7)),
    kind: "range",
    rev: 1,
    src: {
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, newToken());
}

// A press on a pane button, then the pane's work to the end.
export async function press(id: string): Promise<void> {
  document.getElementById(id)!.click();
  await settlePpt();
}

// The Settings tab's own "How links reach PowerPoint" select.
export async function chooseMode(mode: "local" | "relay"): Promise<void> {
  const select = document.getElementById("link-transport") as HTMLSelectElement;
  select.value = mode;
  select.dispatchEvent(new Event("change", { bubbles: true }));
  await settlePpt();
}

function words(element: Element): string {
  return (element.textContent ?? "").replace(/\s+/g, " ").trim();
}

// Hidden by itself, or by an ancestor below `root`. The tab that holds `root`
// may be hidden (one tab shows at a time): that is the tab's business.
function isShown(element: Element, root: Element): boolean {
  for (
    let node: Element | null = element;
    node !== null && node !== root;
    node = node.parentElement
  ) {
    if ((node as HTMLElement).hidden) return false;
  }
  return true;
}

// What a user would read inside `rootSelector`: the text of every `items`
// element (paragraphs by default) that is not hidden, in page order.
export function shownText(rootSelector: string, items = "p"): string[] {
  const root = document.querySelector(rootSelector);
  if (root === null) throw new Error(`no ${rootSelector} on the pane`);
  return [...root.querySelectorAll(items)]
    .filter((item) => isShown(item, root))
    .map(words);
}

// One line of the pane by id: whether its own `hidden` is off, and its words.
export function lineById(id: string): { shown: boolean; text: string } {
  const found = document.getElementById(id);
  if (found === null) throw new Error(`no #${id} on the pane`);
  return { shown: !found.hidden, text: words(found) };
}

// The line the Inbox list shows when nothing waits.
export function emptyInboxText(): string {
  return words(document.querySelector("#inbox-list .hint")!);
}

export function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}
