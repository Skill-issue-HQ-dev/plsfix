// The unused-style scrubber against the strict fake host: which custom styles
// one pass over the workbook calls unused, which sheets it refuses to read, and
// what the delete does with a workbook it could not read in full.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(sheets: string[] = ["Model", "Data"]): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets });
  helpers = host.helpers;
  workbook = host.workbook;
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

// Two custom styles beside Excel's own Normal: one worn by a cell on a hidden
// sheet, one worn by nobody.
function seedStyles(): void {
  helpers.addStyle("Assumption");
  helpers.addStyle("Old header");
  helpers.seed("Data!B2", [["Growth"]]);
  helpers.setStyle("Data!B2", "Assumption");
  helpers.sheet("Data").visibility = "Hidden";
}

// A four-cell block, which is over the cap the scans below run with.
function seedBigSheet(): void {
  helpers.seed("Model!A1", [
    [1, 2],
    [3, 4],
  ]);
}

beforeEach(async () => {
  await boot();
});

describe("list unused styles", () => {
  it("lists the style nobody wears and counts a hidden sheet as use", async () => {
    seedStyles();

    expect(await smt.listUnusedStyles()).toEqual({
      unused: ["Old header"],
      total: 3,
      skippedSheets: [],
    });
  });

  it("never offers a built-in style, however unused", async () => {
    helpers.addStyle("Comma", true);

    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 2,
      skippedSheets: [],
    });
  });

  it("skips a sheet whose used range is past the cap and says which", async () => {
    seedStyles();
    seedBigSheet();

    expect(await smt.listUnusedStyles(3)).toEqual({
      unused: ["Old header"],
      total: 3,
      skippedSheets: ["Model"],
    });
  });

  it("reads an empty workbook without touching a null range", async () => {
    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 1,
      skippedSheets: [],
    });
  });

  // Style properties are the heaviest read of the three workbook scans, and no
  // single sheet has to be large for the request to be: the running total is
  // what stops it, and the sheets left out are named.
  it("stops at the scan cap and names the sheets it did not read", async () => {
    await boot(["Model", "Data", "Notes"]);
    helpers.addStyle("Assumption");
    for (const sheet of ["Model", "Data", "Notes"]) {
      helpers.seed(`${sheet}!A1`, [
        ["a", "b"],
        ["c", "d"],
      ]);
    }
    helpers.setStyle("Notes!A1", "Assumption");

    // Two sheets fit under the total; the third is skipped, so the style only
    // it wears reads as unused and the delete has to refuse.
    expect(await smt.listUnusedStyles(4, 8)).toEqual({
      unused: ["Assumption"],
      total: 2,
      skippedSheets: ["Notes"],
    });
    expect(
      await rejects(() => smt.deleteUnusedStyles(["Assumption"], 4, 8)),
    ).toBe("styles: some sheets were too large to scan");
    expect(workbook.styles).toHaveLength(2);
  });
});

describe("delete unused styles", () => {
  it("deletes the style and leaves the rest of the table alone", async () => {
    seedStyles();

    expect(await smt.deleteUnusedStyles(["Old header"])).toBe(1);
    expect(workbook.styles.map((style) => style.name)).toEqual([
      "Normal",
      "Assumption",
    ]);
  });

  it("leaves a style a cell is still wearing where it is", async () => {
    seedStyles();

    expect(await smt.deleteUnusedStyles(["Assumption", "Old header"])).toBe(1);
    expect(workbook.styles.map((style) => style.name)).toEqual([
      "Normal",
      "Assumption",
    ]);
  });

  it("refuses the whole batch when a sheet was too large to scan", async () => {
    seedStyles();
    seedBigSheet();

    expect(await rejects(() => smt.deleteUnusedStyles(["Old header"], 3))).toBe(
      "styles: some sheets were too large to scan",
    );
    expect(workbook.styles).toHaveLength(3);
  });
});

// The styles pls,fix saved are the user's own work: a freshly saved one that no
// cell wears yet is not "unused", so the scrubber neither lists nor deletes it.
describe("styles pls,fix registered", () => {
  const REGISTRY = "plsfix.styles.v1";
  const register = (...names: string[]): void =>
    helpers.setSetting(
      REGISTRY,
      JSON.stringify({
        v: 1,
        styles: names.map((name) => ({ name, kind: "user" })),
      }),
    );

  it("are not listed as unused, whatever case the list spells them in", async () => {
    seedStyles();
    register("old HEADER");

    expect(await smt.listUnusedStyles()).toEqual({
      unused: [],
      total: 3,
      skippedSheets: [],
    });
  });

  it("are never deleted, even when the caller names them", async () => {
    seedStyles();
    register("Old header");

    expect(await smt.deleteUnusedStyles(["Old header"])).toBe(0);
    expect(workbook.styles.map((style) => style.name)).toEqual([
      "Normal",
      "Assumption",
      "Old header",
    ]);
  });

  it("leave an unregistered unused style as deletable as before", async () => {
    seedStyles();
    helpers.addStyle("Margin %");
    register("Margin %");

    expect((await smt.listUnusedStyles()).unused).toEqual(["Old header"]);
    expect(await smt.deleteUnusedStyles(["Old header", "Margin %"])).toBe(1);
  });

  it("stay exempt when a newer pls,fix wrote the list", async () => {
    seedStyles();
    helpers.setSetting(
      REGISTRY,
      JSON.stringify({ v: 9, styles: [{ name: "Old header", kind: "user" }] }),
    );

    expect((await smt.listUnusedStyles()).unused).toEqual([]);
  });

  it("read a setting that is not a list as no exemption at all", async () => {
    seedStyles();
    helpers.setSetting(REGISTRY, "{ not json");

    expect((await smt.listUnusedStyles()).unused).toEqual(["Old header"]);
  });
});
