// Updating a style from the active cell (every wearer follows), deleting one
// (every wearer goes back to Normal, counted first), and the eight starter
// styles drawn from the palette: against the strict fake host.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { uninstallFakeHost, type FakeHelpers } from "./fakehost";
import {
  bootStyleHost,
  readStyle,
  rejects,
  REGISTRY_KEY,
  storedKinds,
  storedNames,
  styleNames,
  type StyleHost,
} from "./custom-styles.support";
import { starterStyles, STARTER_NAMES } from "../src/custom-styles";
import { tint } from "../src/settings";

let host: StyleHost;
let helpers: FakeHelpers;

beforeEach(async () => {
  host = await bootStyleHost();
  helpers = host.helpers;
});

afterEach(() => {
  uninstallFakeHost();
});

// Saves "Margin %" from a cell with a pale olive fill and puts it on A1:A2.
async function marginOnTwoCells(): Promise<void> {
  helpers.setFill("Model!B2", { color: "#F2E8DF", pattern: "Solid" });
  helpers.setFont("Model!B2", { bold: true, color: "#1F1D1B" });
  helpers.setNumberFormat("Model!B2", "0.0%");
  helpers.setActiveCell("Model!B2");
  await host.smt.saveCustomStyle("Margin %");
  helpers.select("Model!A1:A2");
  await host.smt.applyCustomStyle("Margin %");
}

// A different look in C3, the cell an update is read from.
function newLook(): void {
  helpers.setFill("Model!C3", { color: "#CCE0FF", pattern: "Solid" });
  helpers.setFont("Model!C3", { bold: false, italic: true, color: "#AA0000" });
  helpers.setNumberFormat("Model!C3", "0.00");
  helpers.setActiveCell("Model!C3");
}

describe("update a style from the active cell", () => {
  it("rewrites the style, and every cell wearing it follows", async () => {
    await marginOnTwoCells();
    newLook();

    expect(await host.smt.updateCustomStyle("Margin %")).toBe(
      "Updated Margin %: 2 cells wear it.",
    );

    expect(await readStyle("Margin %")).toMatchObject({
      numberFormat: "0.00",
      fill: { color: "#CCE0FF" },
      font: { bold: false, italic: true, color: "#AA0000" },
    });
    for (const address of ["Model!A1", "Model!A2"]) {
      expect(helpers.cell(address).style, address).toBe("Margin %");
      expect(helpers.fill(address).color, address).toBe("#CCE0FF");
      expect(helpers.numberFormat(address), address).toBe("0.00");
    }
    expect(helpers.fill("Model!A3").pattern).toBe("None");
  });

  it("counts one wearer in the singular and none in plain words", async () => {
    await marginOnTwoCells();
    newLook();
    helpers.setStyle("Model!A2", "Normal");
    expect(await host.smt.updateCustomStyle("Margin %")).toBe(
      "Updated Margin %: 1 cell wears it.",
    );

    helpers.setStyle("Model!A1", "Normal");
    expect(await host.smt.updateCustomStyle("Margin %")).toBe(
      "Updated Margin %: no cell wears it yet.",
    );
  });

  it("leaves the count out when a sheet is too large to scan", async () => {
    await marginOnTwoCells();
    newLook();

    expect(await host.smt.updateCustomStyle("Margin %", 3)).toBe(
      "Updated Margin %.",
    );
    expect((await readStyle("Margin %")).fill.color).toBe("#CCE0FF");
  });

  it("keeps the list as it was: same names, same kinds", async () => {
    await marginOnTwoCells();
    newLook();
    await host.smt.updateCustomStyle("Margin %");

    expect(storedNames(helpers)).toEqual(["Margin %"]);
    expect(storedKinds(helpers)).toEqual(["user"]);
  });

  it("refuses a style Excel no longer has, and forgets it", async () => {
    await marginOnTwoCells();
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Margin %").delete();
      await context.sync();
    });
    newLook();

    expect(await rejects(() => host.smt.updateCustomStyle("Margin %"))).toBe(
      "The style Margin % is no longer in this workbook.",
    );
    expect(storedNames(helpers)).toEqual([]);
  });
});

describe("delete a style", () => {
  it("counts the cells wearing it, whatever the case of the name", async () => {
    await marginOnTwoCells();

    expect(await host.smt.countStyleWearers("margin %")).toBe(2);
    expect(await host.smt.countStyleWearers("Spare")).toBe(0);
  });

  it("refuses to count, and so to delete, when a sheet is too large to scan", async () => {
    await marginOnTwoCells();

    const said = await rejects(() => host.smt.countStyleWearers("Margin %", 3));
    expect(said).toBe(
      "Some sheets are too large to scan (Model), so pls,fix cannot say how many cells wear the style. Delete it in Excel's Cell Styles instead.",
    );
    expect(await rejects(() => host.smt.deleteCustomStyle("Margin %", 3))).toBe(
      said,
    );
    expect(await styleNames()).toContain("Margin %");
    expect(storedNames(helpers)).toEqual(["Margin %"]);
  });

  it("deletes the style and sends every wearer back to Normal", async () => {
    await marginOnTwoCells();

    expect(await host.smt.deleteCustomStyle("Margin %")).toBe(
      "Deleted style Margin %: 2 cells went back to Normal.",
    );

    expect(await styleNames()).not.toContain("Margin %");
    for (const address of ["Model!A1", "Model!A2"]) {
      expect(helpers.cell(address).style, address).toBe("Normal");
      expect(helpers.fill(address).pattern, address).toBe("None");
      expect(helpers.numberFormat(address), address).toBe("General");
    }
    expect(storedNames(helpers)).toEqual([]);
  });

  it("says one cell in the singular, and nothing about cells when no one wore it", async () => {
    await marginOnTwoCells();
    helpers.setStyle("Model!A2", "Normal");
    expect(await host.smt.deleteCustomStyle("Margin %")).toBe(
      "Deleted style Margin %: 1 cell went back to Normal.",
    );

    helpers.setActiveCell("Model!B2");
    await host.smt.saveCustomStyle("Spare");
    expect(await host.smt.deleteCustomStyle("Spare")).toBe(
      "Deleted style Spare.",
    );
  });

  it("leaves the other styles of the list alone", async () => {
    await marginOnTwoCells();
    helpers.setActiveCell("Model!C3");
    await host.smt.saveCustomStyle("Spread");

    await host.smt.deleteCustomStyle("Margin %");

    expect(storedNames(helpers)).toEqual(["Spread"]);
  });

  it("is already done when Excel's own Cell Styles removed the style first", async () => {
    await marginOnTwoCells();
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Margin %").delete();
      await context.sync();
    });

    expect(await host.smt.deleteCustomStyle("Margin %")).toBe(
      "The style Margin % was already gone from this workbook.",
    );
    expect(storedNames(helpers)).toEqual([]);
  });

  it("never touches one of Excel's own styles", async () => {
    expect(await rejects(() => host.smt.deleteCustomStyle("Normal"))).toBe(
      "Normal is one of Excel's own styles and cannot be deleted.",
    );
  });
});

describe("starter styles from the palette", () => {
  const settingsOf = () => host.settings.getActiveSettings();

  it("adds the eight, in the palette's colors, and lists them first", async () => {
    expect(await host.smt.addStarterStyles()).toBe("Added 8 starter styles.");

    const names = await styleNames();
    for (const name of STARTER_NAMES) expect(names, name).toContain(name);
    expect(storedNames(helpers)).toEqual([...STARTER_NAMES]);
    expect(new Set(storedKinds(helpers))).toEqual(new Set(["starter"]));
  });

  it("writes each style as the palette defines it", async () => {
    await host.smt.addStarterStyles();

    for (const spec of starterStyles(settingsOf())) {
      const read = await readStyle(spec.name);
      expect(read.include, spec.name).toEqual({
        number: spec.include.number,
        font: spec.include.font,
        patterns: spec.include.patterns,
        border: spec.include.border,
        alignment: spec.include.alignment,
        protection: false,
      });
      if (spec.include.patterns && spec.fill) {
        expect(read.fill.color, spec.name).toBe(spec.fill);
      }
      if (spec.include.font && spec.font?.color) {
        expect(read.font.color, spec.name).toBe(spec.font.color);
      }
      if (spec.include.font) {
        expect(read.font.name, spec.name).toBe(settingsOf().font);
        expect(read.font.size, spec.name).toBe(spec.font?.size);
      }
    }
    const result = await readStyle("pls,fix Result");
    expect(result.borders.top).toMatchObject({ style: "Double" });
    expect(result.borders.bottom).toMatchObject({ style: "None" });
    const total = await readStyle("pls,fix Total");
    expect(total.borders.top).toMatchObject({
      style: "Continuous",
      weight: "Thin",
    });
    expect(total.borders.bottom).toMatchObject({ style: "Double" });
    const highlight = await readStyle("pls,fix Highlight");
    expect(highlight.fill.color).toBe(tint(settingsOf().accent, 0.8));
  });

  it("is idempotent: a second press restyles the eight and adds none", async () => {
    await host.smt.addStarterStyles();
    const before = await styleNames();

    expect(await host.smt.addStarterStyles()).toBe(
      "Restyled 8 starter styles from the palette.",
    );

    expect(await styleNames()).toEqual(before);
    expect(storedNames(helpers)).toEqual([...STARTER_NAMES]);
  });

  it("restyles every cell wearing a starter after the palette changed", async () => {
    await host.smt.addStarterStyles();
    helpers.select("Model!A1:B1");
    await host.smt.applyCustomStyle("pls,fix Title");
    expect(helpers.fill("Model!A1").color).toBe(settingsOf().primary);

    host.settings.setActiveSettings({
      ...settingsOf(),
      primary: "#AA0000",
      brand: ["#AA0000", settingsOf().accent],
    });
    await host.smt.addStarterStyles();

    expect(helpers.fill("Model!A1").color).toBe("#AA0000");
    expect(helpers.fill("Model!B1").color).toBe("#AA0000");
    expect(helpers.cell("Model!A1").style).toBe("pls,fix Title");
  });

  it("adds what is missing and restyles what exists, and says both", async () => {
    await host.smt.addStarterStyles();
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("pls,fix Note").delete();
      await context.sync();
    });

    expect(await host.smt.addStarterStyles()).toBe(
      "Added 1 and restyled 7 starter styles from the palette.",
    );
    expect(storedNames(helpers)).toEqual([...STARTER_NAMES]);
  });

  it("keeps the user's own styles, after the starters", async () => {
    helpers.setActiveCell("Model!C3");
    await host.smt.saveCustomStyle("Spread");
    await host.smt.addStarterStyles();

    expect(storedNames(helpers)).toEqual([...STARTER_NAMES, "Spread"]);
    expect(storedKinds(helpers).at(-1)).toBe("user");
  });

  // Updated from the pane a starter is the user's own and no longer restyled (custom-styles.starter-
  // update.*); changed in Excel's own Cell Styles the list is not told, so the press still owns it.
  it("puts every starter edited in Excel's own Cell Styles back to its own definition, font and all", async () => {
    await host.smt.addStarterStyles();
    // Every field a style can carry, set to something else.
    await Excel.run(async (context) => {
      for (const name of STARTER_NAMES) {
        const style = context.workbook.styles.getItem(name);
        style.font.name = "Courier New";
        style.font.size = 18;
        style.font.bold = true;
        style.font.italic = true;
        style.font.color = "#AA0000";
        style.fill.color = "#CCE0FF";
        style.numberFormat = "0.00";
        style.includeNumber = true;
        style.includePatterns = true;
        style.includeBorder = true;
        style.includeAlignment = true;
      }
      await context.sync();
    });
    expect((await readStyle("pls,fix Input")).include.patterns).toBe(true);

    await host.smt.addStarterStyles();

    const bold: Record<string, boolean> = {
      "pls,fix Title": true,
      "pls,fix Header": true,
      "pls,fix Result": true,
      "pls,fix Total": true,
    };
    for (const spec of starterStyles(settingsOf())) {
      const read = await readStyle(spec.name);
      expect(read.include, spec.name).toEqual({
        number: false,
        font: spec.include.font,
        patterns: spec.include.patterns,
        border: spec.include.border,
        alignment: false,
        protection: false,
      });
      if (!spec.include.font) continue;
      expect(read.font, spec.name).toEqual({
        name: settingsOf().font,
        size: spec.name === "pls,fix Title" ? 15 : 10,
        bold: bold[spec.name] ?? false,
        italic: spec.name === "pls,fix Note",
        color: spec.font?.color,
      });
    }
  });

  it("refuses to write over a list a newer pls,fix saved", async () => {
    helpers.setSetting(REGISTRY_KEY, JSON.stringify({ v: 2, styles: [] }));
    const before = await styleNames();

    expect(await rejects(() => host.smt.addStarterStyles())).toMatch(
      /newer pls,fix\. Update the add-in\.$/,
    );
    expect(await styleNames()).toEqual(before);
  });
});
