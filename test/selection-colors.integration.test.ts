// readSelectionFills over the strict fake host: the distinct fills of the
// selected cells in reading order, one read of the fills, nothing written and
// nothing spent on pls,fix Undo, the selection cap honoured.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(web = false): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model", "Data"], web }).helpers;
  smt = await import("../src/excel");
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

function solid(color: string): { color: string; pattern: string } {
  return { color, pattern: "Solid" };
}

beforeEach(async () => {
  await boot();
});

describe("readSelectionFills", () => {
  it("lists each distinct fill once, in reading order", async () => {
    helpers.setFill("Model!A1", solid(GOLD));
    helpers.setFill("Model!B1", solid(MARBLE));
    helpers.setFill("Model!A2", solid(MARBLE));
    helpers.setFill("Model!B2", solid(GREEN));
    helpers.select("Model!A1:B2");
    expect(await smt.readSelectionFills()).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("reads row by row, not column by column", async () => {
    helpers.setFill("Model!A1", solid(GREEN));
    helpers.setFill("Model!B1", solid(GOLD));
    helpers.setFill("Model!A2", solid(MARBLE));
    helpers.select("Model!A1:B2");
    expect(await smt.readSelectionFills()).toEqual([GREEN, GOLD, MARBLE]);
  });

  it("skips cells with no fill", async () => {
    helpers.setFill("Model!B2", solid(GOLD));
    helpers.select("Model!A1:C3");
    expect(await smt.readSelectionFills()).toEqual([GOLD]);
  });

  it("answers an empty list for a selection with no fill at all", async () => {
    helpers.select("Model!A1:C3");
    expect(await smt.readSelectionFills()).toEqual([]);
  });

  it("reads every area of a ctrl-clicked selection, the first area first", async () => {
    helpers.setFill("Model!D4", solid(MARBLE));
    helpers.setFill("Model!A1", solid(GOLD));
    helpers.selectAreas(["Model!D4", "Model!A1"]);
    expect(await smt.readSelectionFills()).toEqual([MARBLE, GOLD]);
  });

  it("reads a solid fill on Excel for the web, which names no pattern", async () => {
    await boot(true);
    helpers.setFill("Model!A1", solid(GOLD));
    helpers.select("Model!A1:B1");
    expect(await smt.readSelectionFills()).toEqual([GOLD]);
  });

  it("is a read: nothing written, nothing on the Undo stack", async () => {
    helpers.setFill("Model!A1", solid(GOLD));
    helpers.select("Model!A1:B1");
    const before = helpers.cellMap("Model");
    await smt.readSelectionFills();
    expect(helpers.cellMap("Model")).toEqual(before);
    expect(smt.undoTarget()).toBeNull();
  });

  it("reads the fills in one sync beyond the selection's own", async () => {
    helpers.setFill("Model!A1", solid(GOLD));
    helpers.select("Model!A1:C3");
    const before = helpers.syncCount();
    await smt.readSelectionFills();
    // The areas, their cell count against the cap, then the one fills read.
    expect(helpers.syncCount() - before).toBeLessThanOrEqual(3);
  });

  it("refuses a selection past the cap before reading a single fill", async () => {
    helpers.select("Model!A1:A5001");
    expect(await rejects(() => smt.readSelectionFills())).toBe(
      "Read from selection supports up to 5,000 selected cells at once.",
    );
  });
});
