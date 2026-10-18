// The Excel Links tab as a modeller drives it, over the REAL adapters: the
// shipped taskpane.html, the strict fake host and a recorded fake relay, so a
// suite can say which requests a press sent (or that none left the pane). Only
// the guard and the toast are the rig's own. Import from a jsdom suite.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import {
  saveTransport,
  type LinkTransport,
} from "../src/link/transport-setting";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import type { Guard } from "../src/ui/guard";
import type { ToastKind } from "../src/ui/toast";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import { recordCalls } from "./relay-calls";

export interface TabRig {
  // What the tab was handed, recording every request it makes.
  relay: FakeRelay;
  // The same relay, unrecorded: seed it and read its state without a trace.
  fake: FakeRelay;
  // Every relay method the tab called, in order.
  calls: string[];
  // Every global fetch the pane made on its own: the relay is a fake here, so a
  // direct fetch would slip past the recorder above (and is refused).
  fetches: string[];
  helpers: FakeHelpers;
  links: typeof LinksModule;
  keyStore: KeyStore;
  // The paired link key, or null on a device that never made one.
  ws: Workspace | null;
  messages: string[];
  errors: string[];
  toasts: { message: string; kind?: ToastKind; details?: string }[];
  // Installs the tab over the workbook as it stands and waits for its boot.
  mount(): Promise<void>;
  click(id: string): void;
  // Ticks the list row of one link (left ticked if it already is).
  tick(id: string): void;
  settle(): Promise<void>;
}

const SETTLE_TURNS = 25;

// The afterEach of every suite that opens a workbook here.
export function closeWorkbook(): void {
  uninstallFakeHost();
  vi.unstubAllGlobals();
}

export function memoryStore(): KeyStore {
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

// A workbook with a 2 x 5 block at Model!B4 selected, on a device set to
// `transport`, with or without a link key. The tab is not mounted yet, so a
// test can make links first the way an earlier session would have.
export async function openWorkbook(options: {
  transport: LinkTransport;
  key: boolean;
}): Promise<TabRig> {
  enableStrictLoadSemantics();
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  host.workbook.fileUrl = "/Volumes/Work/Models/Model_v4.xlsx";
  host.helpers.seed("Model!B4", [
    [1, 2, 3, 4, 5],
    [6, 7, 8, 9, 10],
  ]);
  host.helpers.select("Model!B4:F5");
  const keyStore = memoryStore();
  await saveTransport(keyStore, options.transport);
  const ws = options.key ? await createWorkspace(keyStore) : null;
  const fake = new FakeRelay();
  const { relay, calls } = recordCalls(fake);
  const links = await import("../src/excel/links");
  const fetches: string[] = [];
  vi.stubGlobal("fetch", (input: unknown) => {
    fetches.push(String(input));
    return Promise.reject(new Error("no network in this rig"));
  });

  const pending: Promise<void>[] = [];
  const messages: string[] = [];
  const errors: string[] = [];
  const toasts: TabRig["toasts"] = [];
  const guard: Guard = (run) => {
    const done = (async () => {
      try {
        messages.push(await run());
      } catch (error) {
        errors.push(error instanceof Error ? error.message : String(error));
      }
    })();
    pending.push(done);
    return done;
  };

  async function settle(): Promise<void> {
    for (let turn = 0; turn < SETTLE_TURNS; turn += 1) {
      while (pending.length > 0) await Promise.all(pending.splice(0));
      await new Promise((done) => setTimeout(done, 0));
    }
  }

  return {
    relay,
    fake,
    calls,
    fetches,
    helpers: host.helpers,
    links,
    keyStore,
    ws,
    messages,
    errors,
    toasts,
    settle,
    async mount() {
      document.body.innerHTML = readFileSync(
        join(process.cwd(), "taskpane.html"),
        "utf8",
      );
      const { installLinksTab } = await import("../src/pane/links-tab");
      installLinksTab({
        guard,
        toast: {
          show: (message, kind, details) => {
            toasts.push({ message, kind, details });
          },
        },
        relay,
        keyStore,
        root: document,
      });
      await settle();
    },
    click(id) {
      document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
    },
    tick(id) {
      const box = document.querySelector<HTMLInputElement>(
        `tr[data-link-id="${id}"] input[type="checkbox"]`,
      );
      if (box === null) throw new Error(`no tick box for link ${id}`);
      if (box.checked) return;
      box.click();
    },
  };
}
