// The football field data block through the REAL Football field button, over
// the strict fake host: insert the block, leave it selected as the insert does,
// press the chart's button and get the chart - the whole promise of the block,
// proven on the production path rather than a restatement of its rules. The
// fake host has no formula engine, so the Midpoint column stays empty here;
// the chart reads label, low and high only, which is the point.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  type FakeWorkbook,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model"] });
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

beforeEach(async () => {
  await boot();
});

describe("football field data, then the Football field button", () => {
  it("leaves the block selected, and that selection is the chart's input", async () => {
    helpers.select("Model!B2");
    await smt.insertTemplate("football-field");
    expect(workbook.selectionAddress()).toBe("Model!B2:E7");

    expect(await smt.insertFootballField()).toBe(
      "Football field added: 5 ranges; only label, low and high are used",
    );
  });

  it("builds a stacked bar over a helper block of the five methods beside the data", async () => {
    helpers.select("Model!B2");
    await smt.insertTemplate("football-field");
    await smt.insertFootballField();

    expect(
      ["F2", "G2", "H2"].map((at) => helpers.value(`Model!${at}`)),
    ).toEqual(["Method", "Low", "Range"]);
    const rows = [3, 4, 5, 6, 7].map((row) =>
      ["F", "G", "H"].map((column) => helpers.value(`Model!${column}${row}`)),
    );
    expect(rows).toEqual([
      ["Trading comps", 465, 90],
      ["Precedents", 400, 200],
      ["DCF", 450, 130],
      ["52-week range", 410, 100],
      ["Analyst targets", 470, 90],
    ]);

    expect(workbook.charts).toHaveLength(1);
    expect(workbook.charts[0]).toMatchObject({
      chartType: "BarStacked",
      seriesBy: "Columns",
      sourceAddress: "Model!F2:H7",
      title: "Valuation range",
    });
  });

  it("works wherever the block lands, with room for the helper beside it", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate("football-field");
    await smt.insertFootballField();
    expect(workbook.charts[0]?.sourceAddress).toBe("Model!E1:G6");
    expect(helpers.value("Model!E2")).toBe("Trading comps");
  });

  it("refuses a second press on the same block rather than drawing twice", async () => {
    helpers.select("Model!B2");
    await smt.insertTemplate("football-field");
    await smt.insertFootballField();
    helpers.select("Model!B2:E7");

    expect(await rejects(() => smt.insertFootballField())).toMatch(
      /Football field/,
    );
    expect(workbook.charts).toHaveLength(1);
  });
});
