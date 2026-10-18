// Saving the look of a cell as a real Excel cell style, putting it on the
// selection, and listing what the workbook has: against the strict fake host.
// Update, delete and the starter styles are in custom-styles.manage.*.

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

let host: StyleHost;
let helpers: FakeHelpers;

// A source cell wearing something on every field a style carries.
function dress(address: string): void {
  helpers.setFont(address, {
    name: "Aptos",
    size: 12,
    bold: true,
    italic: true,
    color: "#1F1D1B",
  });
  helpers.setFill(address, { color: "#F2E8DF", pattern: "Solid" });
  helpers.setNumberFormat(address, "#,##0.0");
  helpers.setAlignment(address, "Right");
  const [sheet, local] = address.split("!") as [string, string];
  const rect = {
    row: Number(local.slice(1)) - 1,
    col: local.charCodeAt(0) - 65,
  };
  helpers.sheet(sheet).edit(rect.row, rect.col).borders.bottom = {
    style: "Continuous",
    color: "#B27E54",
    weight: "Medium",
  };
}

async function save(name: string, from = "Model!B2"): Promise<string> {
  dress(from);
  helpers.setActiveCell(from);
  return host.smt.saveCustomStyle(name);
}

beforeEach(async () => {
  host = await bootStyleHost();
  helpers = host.helpers;
});

afterEach(() => {
  uninstallFakeHost();
});

describe("save a style from the active cell", () => {
  it("adds a real Excel style that carries the cell's whole look", async () => {
    expect(await save("Margin %")).toBe("Saved style Margin %");

    expect(await readStyle("Margin %")).toMatchObject({
      name: "Margin %",
      numberFormat: "#,##0.0",
      horizontalAlignment: "Right",
      fill: { color: "#F2E8DF", pattern: "Solid" },
      font: {
        name: "Aptos",
        size: 12,
        bold: true,
        italic: true,
        color: "#1F1D1B",
      },
      borders: {
        bottom: { style: "Continuous", color: "#B27E54", weight: "Medium" },
        top: { style: "None" },
        left: { style: "None" },
        right: { style: "None" },
      },
    });
  });

  it("writes a None edge as its style alone and leaves a double rule its own weight", async () => {
    dress("Model!B2");
    const cell = helpers.sheet("Model").edit(1, 1);
    // A leftover weight and color under an edge that draws nothing, and a double
    // bottom rule whose weight Excel reads as Thick.
    cell.borders.top = { style: "None", color: "#FF0000", weight: "Thick" };
    cell.borders.bottom = {
      style: "Double",
      color: "#B27E54",
      weight: "Thick",
    };
    helpers.setActiveCell("Model!B2");

    await host.smt.saveCustomStyle("Rules");

    const { borders } = await readStyle("Rules");
    expect(borders.top).toEqual({
      style: "None",
      color: "#000000",
      weight: "Thin",
    });
    expect(borders.bottom).toEqual({
      style: "Double",
      color: "#B27E54",
      weight: "Thin",
    });
  });

  it("includes every category but protection", async () => {
    await save("Margin %");
    expect((await readStyle("Margin %")).include).toEqual({
      number: true,
      font: true,
      patterns: true,
      border: true,
      alignment: true,
      protection: false,
    });
  });

  it("remembers it in the workbook's own list, newest last", async () => {
    await save("Margin %");
    await save("Spread", "Model!C2");

    expect(storedNames(helpers)).toEqual(["Margin %", "Spread"]);
    expect(storedKinds(helpers)).toEqual(["user", "user"]);
  });

  it("saves a cell with no fill as a style with no fill", async () => {
    helpers.setActiveCell("Model!C3");
    await host.smt.saveCustomStyle("Plain");

    expect((await readStyle("Plain")).fill.pattern).toBe("None");
  });

  it("trims the name", async () => {
    expect(await save("  Spread  ")).toBe("Saved style Spread");
    expect(storedNames(helpers)).toEqual(["Spread"]);
  });

  it("is not offered as unused by the style scrubber while no cell wears it", async () => {
    await save("Margin %");

    expect((await host.smt.listUnusedStyles()).unused).toEqual([]);
  });

  it("refuses a name that is taken, built-in styles included, and writes nothing", async () => {
    await save("Margin %");
    const before = await styleNames();

    for (const [typed, said] of [
      ["margin %", "A style named margin % already exists in this workbook."],
      ["Normal", "A style named Normal already exists in this workbook."],
    ] as const) {
      expect(await rejects(() => host.smt.saveCustomStyle(typed))).toBe(said);
    }
    expect(await styleNames()).toEqual(before);
    expect(storedNames(helpers)).toEqual(["Margin %"]);
  });

  it("refuses an empty and a too long name, in one sentence each", async () => {
    expect(await rejects(() => host.smt.saveCustomStyle("   "))).toBe(
      "Type a name for the style first.",
    );
    expect(await rejects(() => host.smt.saveCustomStyle("x".repeat(41)))).toBe(
      "A style name can be up to 40 characters.",
    );
    expect(helpers.setting(REGISTRY_KEY)).toBeNull();
  });

  it("refuses to write over a list a newer pls,fix saved", async () => {
    helpers.setSetting(REGISTRY_KEY, JSON.stringify({ v: 2, styles: [] }));
    const before = await styleNames();

    expect(await rejects(() => host.smt.saveCustomStyle("Spread"))).toMatch(
      /newer pls,fix\. Update the add-in\.$/,
    );
    expect(await styleNames()).toEqual(before);
    expect(JSON.parse(helpers.setting(REGISTRY_KEY) ?? "{}")).toEqual({
      v: 2,
      styles: [],
    });
  });
});

describe("save on Excel for the web", () => {
  it("reads a solid fill the web answers with no pattern, and an unfilled cell as none", async () => {
    uninstallFakeHost();
    host = await bootStyleHost(true);
    helpers = host.helpers;

    await save("Margin %");
    expect((await readStyle("Margin %")).fill).toMatchObject({
      color: "#F2E8DF",
      pattern: "Solid",
    });

    helpers.setActiveCell("Model!C3");
    await host.smt.saveCustomStyle("Plain");
    expect((await readStyle("Plain")).fill.pattern).toBe("None");
    expect((await readStyle("Plain")).borders.bottom).toMatchObject({
      style: "None",
    });
  });
});

describe("apply a style", () => {
  beforeEach(async () => {
    await save("Margin %");
    helpers.select("Model!A1:A3");
  });

  it("puts the style on every selected cell, with its look", async () => {
    expect(await host.smt.applyCustomStyle("Margin %")).toBe("Style: Margin %");

    for (const address of ["Model!A1", "Model!A2", "Model!A3"]) {
      const cell = helpers.cell(address);
      expect(cell.style, address).toBe("Margin %");
      expect(cell.fill, address).toMatchObject({
        color: "#F2E8DF",
        pattern: "Solid",
      });
      expect(cell.font, address).toMatchObject({ bold: true, name: "Aptos" });
      expect(cell.numberFormat, address).toBe("#,##0.0");
    }
    expect(helpers.cell("Model!B1").style).toBe("Normal");
  });

  it("takes the style's own spelling for the receipt", async () => {
    expect(await host.smt.applyCustomStyle("margin %")).toBe("Style: Margin %");
  });

  it("covers every area of a ctrl-clicked selection", async () => {
    helpers.selectAreas(["Model!A1:A2", "Model!C3"]);
    await host.smt.applyCustomStyle("Margin %");

    for (const address of ["Model!A1", "Model!A2", "Model!C3"]) {
      expect(helpers.cell(address).style, address).toBe("Margin %");
    }
    expect(helpers.cell("Model!A3").style).toBe("Normal");
  });

  it("is taken back by pls,fix Undo: the old style and the cells' own look", async () => {
    helpers.setFont("Model!A1", { bold: false, color: "#AA0000" });
    helpers.setNumberFormat("Model!A1", "0.00");
    await host.smt.applyCustomStyle("Margin %");

    await host.smt.undoLastAction();

    expect(helpers.cell("Model!A1").style).toBe("Normal");
    expect(helpers.font("Model!A1")).toMatchObject({
      bold: false,
      color: "#AA0000",
    });
    expect(helpers.numberFormat("Model!A1")).toBe("0.00");
    expect(helpers.fill("Model!A1").pattern).toBe("None");
  });

  it("refuses a style Excel's own Cell Styles no longer has, and forgets it", async () => {
    await save("Spread", "Model!C2");
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Spread").delete();
      await context.sync();
    });

    expect(await rejects(() => host.smt.applyCustomStyle("Spread"))).toBe(
      "The style Spread is no longer in this workbook.",
    );
    expect(storedNames(helpers)).toEqual(["Margin %"]);
    expect(helpers.cell("Model!A1").style).toBe("Normal");
  });

  it("answers a protected sheet in words and changes nothing", async () => {
    helpers.protectSheet("Model", []);

    expect(await rejects(() => host.smt.applyCustomStyle("Margin %"))).toBe(
      "Style: this sheet is protected, nothing was changed",
    );
    expect(helpers.cell("Model!A1").style).toBe("Normal");
  });
});

describe("list the styles", () => {
  it("is empty for a workbook with no pls,fix styles", async () => {
    expect(await host.smt.listCustomStyles()).toEqual({ styles: [] });
  });

  it("gives what a row needs to draw its preview", async () => {
    await save("Margin %");

    expect(await host.smt.listCustomStyles()).toEqual({
      styles: [
        {
          name: "Margin %",
          kind: "user",
          fill: "#F2E8DF",
          fontColor: "#1F1D1B",
          bold: true,
        },
      ],
    });
  });

  it("shows no fill for a style with none, and no font for one that leaves the font alone", async () => {
    helpers.setActiveCell("Model!C3");
    await host.smt.saveCustomStyle("Plain");
    await Excel.run(async (context) => {
      context.workbook.styles.add("Hue");
      const style = context.workbook.styles.getItem("Hue");
      style.fill.color = "#FF0000";
      style.font.bold = true;
      style.includeFont = false;
      await context.sync();
    });
    helpers.setSetting(
      REGISTRY_KEY,
      JSON.stringify({
        v: 1,
        styles: [
          { name: "Plain", kind: "user" },
          { name: "Hue", kind: "starter" },
        ],
      }),
    );

    const { styles } = await host.smt.listCustomStyles();
    expect(styles.find((style) => style.name === "Plain")?.fill).toBeNull();
    expect(styles.find((style) => style.name === "Hue")).toMatchObject({
      kind: "starter",
      fill: "#FF0000",
      fontColor: null,
      bold: false,
    });
  });

  it("leaves out a style deleted through Excel, without writing the list", async () => {
    await save("Margin %");
    await save("Spread", "Model!C2");
    const stored = helpers.setting(REGISTRY_KEY);
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Spread").delete();
      await context.sync();
    });

    const { styles } = await host.smt.listCustomStyles();

    expect(styles.map((style) => style.name)).toEqual(["Margin %"]);
    expect(helpers.setting(REGISTRY_KEY)).toBe(stored);
  });

  it("says so when a newer pls,fix wrote the list", async () => {
    helpers.setSetting(
      REGISTRY_KEY,
      JSON.stringify({ v: 3, styles: [{ name: "Future", kind: "user" }] }),
    );

    expect(await host.smt.listCustomStyles()).toEqual({
      styles: [],
      note: "The styles in this workbook were saved by a newer pls,fix. Update the add-in.",
    });
  });
});
