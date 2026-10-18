// @vitest-environment jsdom
// The native table export against the strict Excel fake: the anchor, the
// registry entry and the inbox note are a range export's, and the payload
// carries the text, the formats and the column widths of the selected cells. A
// cell wearing Excel's defaults carries nothing but its text, and a push
// re-renders through the anchor. The last block drives the Links tab's own
// buttons over the same adapters: Export selection is the table, a picture
// only comes from Export as picture, and a selection over the cap is refused.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinkHighlightModule from "../src/excel/link-highlight";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  anchorName,
  decodeInboxItem,
  decodePayload,
  decodeRegistry,
  REGISTRY_SETTING,
  type Payload,
  type TablePayload,
} from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";
import { closeWorkbook, openWorkbook, type TabRig } from "./links-tab.support";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let ws: Workspace;

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

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Volumes/Work/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../src/excel/links");
  helpers.seed("Model!B4", [
    ["Revenue", 1000],
    ["Costs", -400],
  ]);
  helpers.setFont("Model!B4", { bold: true });
  helpers.setFont("Model!C5", { italic: true, color: "#FF0000", size: 9 });
  // Solid, not just a colour: setFill assigns the raw fill fields, and a real
  // Excel cell can never carry a colour with no pattern (fill.color's own
  // setter always forces Solid) - fillKey needs the pair to say so.
  helpers.setFill("Model!C4", { color: "#EEEEEE", pattern: "Solid" });
  helpers.setAlignment("Model!C4", "Right");
  helpers.sheet("Model").columnWidths.set(1, 96);
  helpers.sheet("Model").columnWidths.set(2, 48);
  helpers.select("Model!B4:C5");
});
afterEach(() => uninstallFakeHost());

// Looked up by id, not just links[0]: a test that exports more than one link
// in the same workbook still finds the right token for each.
function token(id: string): string {
  const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
  const entry = (registry.links as { id: string; token: string }[]).find(
    (one) => one.id === id,
  );
  if (!entry) throw new Error(`${id}: no registry entry`);
  return entry.token;
}

async function tableOf(id: string): Promise<TablePayload> {
  const keys = await deriveLinkKeys(token(id));
  const stored = relay.links.get(id)!;
  const payload = decodePayload(await open(keys.enc, id, stored.blob));
  if (payload.kind !== "table") throw new Error(`${id}: not a table`);
  return payload;
}

describe("exportSelectionAsTable", () => {
  // A host that refuses the rich grid still answers for text and widths: the
  // table lands as a plain one rather than failing before the relay, and the
  // anchor the rich batch bound stays bound.
  it("falls back to the plain grid when the cell-properties read is refused", async () => {
    helpers.failNextCellProperties();
    const result = await links.exportSelectionAsTable(ws, relay);
    const payload = await tableOf(result.id);
    expect(payload.cells).toEqual([
      [{ t: "Revenue" }, { t: "1000", a: "r" }],
      [{ t: "Costs" }, { t: "-400", a: "r" }],
    ]);
    expect(payload.widths).toEqual([96, 48]);
    expect(
      workbook.names.find((n) => n.name === anchorName(result.id)),
    ).toBeDefined();
    expect(
      JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links,
    ).toHaveLength(1);
  });

  it("costs a working host no extra round trip", async () => {
    const before = helpers.syncCount();
    await links.exportSelectionAsTable(ws, relay);
    expect(helpers.syncCount() - before).toBe(6);
  });

  it("anchors the range, records a table entry and pushes the cells, formats and widths", async () => {
    const result = await links.exportSelectionAsTable(ws, relay);
    expect(result.label).toBe("Model!B4:C5 table");
    const name = workbook.names.find((n) => n.name === anchorName(result.id))!;
    expect(name.visible).toBe(false);
    expect(name.formula).toBe("=Model!$B$4:$C$5");
    const registry = JSON.parse(String(helpers.setting(REGISTRY_SETTING)));
    expect(registry.links[0]).toMatchObject({
      id: result.id,
      kind: "table",
      anchor: anchorName(result.id),
      rev: 1,
    });

    const payload = await tableOf(result.id);
    expect(payload).toMatchObject({ kind: "table", rows: 2, cols: 2 });
    expect(payload.widths).toEqual([96, 48]);
    expect(payload.src).toEqual({
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:C5",
      anchor: anchorName(result.id),
    });
    // Text as Excel displays it, and only what differs from the defaults.
    expect(payload.cells).toEqual([
      [
        { t: "Revenue", b: true, z: 11 },
        { t: "1000", f: "#EEEEEE", a: "r", z: 11 },
      ],
      // A number under General alignment says "right" itself.
      [
        { t: "Costs", z: 11 },
        { t: "-400", i: true, c: "#FF0000", z: 9, a: "r" },
      ],
    ]);

    const inbox = await relay.listInbox(ws.id, ws.auth);
    expect(
      decodeInboxItem(await open(ws.enc, ws.id, inbox[0]!.blob)),
    ).toMatchObject({ kind: "table", label: "Model!B4:C5 table" });
  });

  it("re-renders through the anchor on a push, not through the address", async () => {
    const { id } = await links.exportSelectionAsTable(ws, relay);
    helpers.seed("Model!B10", [
      ["Profit", 600],
      ["Margin", "60%"],
    ]);
    helpers.setNameFormula(anchorName(id), "=Model!$B$10:$C$11");

    expect(await links.pushLinks("all", relay)).toEqual({
      pushed: 1,
      missing: 0,
      failed: 0,
      failures: [],
    });
    const payload = await tableOf(id);
    expect(payload.src.ref).toBe("B10:C11");
    expect(payload.cells.map((row) => row.map((cell) => cell.t))).toEqual([
      ["Profit", "600"],
      ["Margin", "60%"],
    ]);
    expect(relay.links.get(id)!.rev).toBe(2);
  });

  // The cap is a table's own, so it is refused before the hidden name is bound.
  it("refuses a selection past 60 rows or 20 columns and anchors nothing", async () => {
    helpers.select("Model!A1:U60");
    await expect(links.exportSelectionAsTable(ws, relay)).rejects.toThrow(
      "Too big for a native table (60 x 20). Select less, or press Export as picture.",
    );
    expect(workbook.names).toEqual([]);
    expect(helpers.setting(REGISTRY_SETTING)).toBeNull();
    expect(relay.links.size).toBe(0);
  });

  // The audit overlay and the linked-cell highlight both paint pattern fills
  // over cells and remember the originals in a FillStore; a table export must
  // read those originals back, never the tint, or a deck picks up pink cells
  // no model ever had.
  it("substitutes the true fill for a cell an overlay is tinting", async () => {
    const highlight: typeof LinkHighlightModule =
      await import("../src/excel/link-highlight");
    const baseline = await links.exportSelectionAsTable(ws, relay);
    const baselineCells = (await tableOf(baseline.id)).cells;

    expect(await highlight.toggleLinkHighlight()).toBe(true);
    helpers.select("Model!B4:C5");
    const tinted = await links.exportSelectionAsTable(ws, relay);
    expect((await tableOf(tinted.id)).cells).toEqual(baselineCells);

    expect(await highlight.toggleLinkHighlight()).toBe(false);
    helpers.select("Model!B4:C5");
    const restored = await links.exportSelectionAsTable(ws, relay);
    expect((await tableOf(restored.id)).cells).toEqual(baselineCells);
  });

  it("lists and removes a table link like any other", async () => {
    const { id } = await links.exportSelectionAsTable(ws, relay);
    const rows = await links.listWorkbookLinks();
    expect(rows).toHaveLength(1);
    expect(rows[0]!.entry.kind).toBe("table");
    expect(rows[0]!.source).toBe("ok");

    await links.goToSource(id);
    expect(workbook.selectionAddress()).toBe("Model!B4:C5");

    await links.removeLink(id, relay);
    expect(
      workbook.names.find((n) => n.name === anchorName(id)),
    ).toBeUndefined();
    expect(relay.links.has(id)).toBe(false);
  });
});

// The Links tab's own buttons over the same real adapters, in relay mode with a
// paired key: the default export is a native table, a picture is a choice made
// with its own button, and the table cap is refused before anything is sent.
describe("the Links tab's export buttons", () => {
  afterEach(() => closeWorkbook());

  async function pairedTab(): Promise<TabRig> {
    const rig = await openWorkbook({ transport: "relay", key: true });
    await rig.mount();
    return rig;
  }

  function entries(rig: TabRig) {
    return decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links;
  }

  async function sentPayload(rig: TabRig, id: string): Promise<Payload> {
    const entry = entries(rig).find((one) => one.id === id)!;
    const keys = await deriveLinkKeys(entry.token);
    const stored = rig.fake.links.get(id)!;
    return decodePayload(await open(keys.enc, id, stored.blob));
  }

  it("Export selection sends a native table, not a picture", async () => {
    const rig = await pairedTab();

    rig.click("export-selection");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages).toEqual(["Sent to PowerPoint: Model!B4:F5 table"]);
    const [entry, ...rest] = entries(rig);
    expect(rest).toEqual([]);
    expect(entry).toMatchObject({ kind: "table", label: "Model!B4:F5 table" });
    const payload = await sentPayload(rig, entry!.id);
    expect(payload).toMatchObject({ kind: "table", rows: 2, cols: 5 });
  });

  it("Export as picture is the one button that sends a picture", async () => {
    const rig = await pairedTab();

    rig.click("export-picture");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages).toEqual(["Sent to PowerPoint: Model!B4:F5"]);
    const [entry, ...rest] = entries(rig);
    expect(rest).toEqual([]);
    expect(entry).toMatchObject({ kind: "range", label: "Model!B4:F5" });
    expect((await sentPayload(rig, entry!.id)).kind).toBe("picture");
  });

  it("keeps a table and a picture of the same cells apart in the list", async () => {
    const rig = await pairedTab();

    rig.click("export-selection");
    await rig.settle();
    rig.click("export-picture");
    await rig.settle();

    const kinds = [
      ...document.querySelectorAll("#workbook-links tr[data-link-id] small"),
    ]
      .map((line) => line.textContent)
      .sort();
    expect(kinds).toEqual([
      "Picture (your choice) · Model!B4:F5",
      "Table · Model!B4:F5",
    ]);
  });

  it("refuses a selection past 60 rows or 20 columns with the table sentence and sends nothing", async () => {
    const rig = await pairedTab();
    rig.helpers.select("Model!A1:U60");

    rig.click("export-selection");
    await rig.settle();

    expect(rig.errors).toEqual([
      "Too big for a native table (60 x 20). Select less, or press Export as picture.",
    ]);
    expect(rig.messages).toEqual([]);
    expect(entries(rig)).toEqual([]);
    expect(rig.fake.links.size).toBe(0);
    expect(rig.calls).not.toContain("putLink");
    expect(rig.calls).not.toContain("postInbox");
  });

  it("still lets the same big selection go as a picture, by choice", async () => {
    const rig = await pairedTab();
    rig.helpers.select("Model!A1:U60");

    rig.click("export-picture");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(entries(rig)).toHaveLength(1);
    expect(entries(rig)[0]).toMatchObject({ kind: "range" });
  });
});
