// How big a workbook Update and Delete can still count the wearers of a style in:
// the cap Find and the model check read under (200 000 cells, per sheet and in
// all), not the style scrubber's 5 000 per sheet and 20 000 in all. Past it the
// answer is in words, naming the sheets. Against the strict fake host.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { uninstallFakeHost, type FakeHelpers } from "./fakehost";
import {
  bootStyleHost,
  rejects,
  styleNames,
  type StyleHost,
} from "./custom-styles.support";

let host: StyleHost;
let helpers: FakeHelpers;

beforeEach(async () => {
  host = await bootStyleHost();
  helpers = host.helpers;
  // "Margin %" worn by Model!A1:A2.
  helpers.setFill("Model!B2", { color: "#F2E8DF", pattern: "Solid" });
  helpers.setActiveCell("Model!B2");
  await host.smt.saveCustomStyle("Margin %");
  helpers.select("Model!A1:A2");
  await host.smt.applyCustomStyle("Margin %");
});

afterEach(() => {
  uninstallFakeHost();
});

// A used range of columns x rows cells: a value in each corner.
function sheetOf(sheet: string, columns: string, rows: number): void {
  helpers.seed(`${sheet}!A1`, [["top"]]);
  helpers.seed(`${sheet}!${columns}${String(rows)}`, [["end"]]);
}

const TOO_LARGE = (sheets: string): string =>
  `Some sheets are too large to scan (${sheets}), so pls,fix cannot say how many cells wear the style. Delete it in Excel's Cell Styles instead.`;

describe("a sheet past the style scrubber's 5 000 cells", () => {
  // AA400: 27 columns x 400 rows = 10 800 cells.
  beforeEach(() => {
    sheetOf("Model", "AA", 400);
  });

  it("is still counted, for the Delete button's count", async () => {
    expect(await host.smt.countStyleWearers("Margin %")).toBe(2);
  });

  it("is still deleted, counted, with every wearer sent back to Normal", async () => {
    expect(await host.smt.deleteCustomStyle("Margin %")).toBe(
      "Deleted style Margin %: 2 cells went back to Normal.",
    );

    expect(await styleNames()).not.toContain("Margin %");
    expect(helpers.cell("Model!A1").style).toBe("Normal");
  });

  it("still gets its wearers counted in the Update receipt", async () => {
    expect(await host.smt.updateCustomStyle("Margin %")).toBe(
      "Updated Margin %: 2 cells wear it.",
    );
  });
});

describe("a workbook past the style scrubber's 20 000 cells in all", () => {
  it("is counted whole: five sheets of 4 800 cells", async () => {
    sheetOf("Model", "T", 240);
    for (const name of ["Data", "Notes", "Checks", "Archive"]) {
      if (name !== "Data") helpers.addSheet(name);
      sheetOf(name, "T", 240);
    }

    expect(await host.smt.countStyleWearers("Margin %")).toBe(2);
  });
});

describe("a sheet past 200 000 cells", () => {
  // Z10000: 26 columns x 10 000 rows = 260 000 cells.
  beforeEach(() => {
    sheetOf("Model", "Z", 10_000);
  });

  it("refuses the count in words and names the sheet", async () => {
    expect(await rejects(() => host.smt.countStyleWearers("Margin %"))).toBe(
      TOO_LARGE("Model"),
    );
  });

  it("refuses the delete the same way, and deletes nothing", async () => {
    expect(await rejects(() => host.smt.deleteCustomStyle("Margin %"))).toBe(
      TOO_LARGE("Model"),
    );
    expect(await styleNames()).toContain("Margin %");
  });

  it("updates the style all the same, without saying how many wear it", async () => {
    expect(await host.smt.updateCustomStyle("Margin %")).toBe(
      "Updated Margin %.",
    );
  });
});
