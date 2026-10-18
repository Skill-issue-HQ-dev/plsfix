// The style-cycle key against the strict fake host: one press puts the next of the workbook's
// pls,fix styles on the selection (list order, then Excel's Normal, then the first), read off the
// active cell, with the apply write path (Undo, every area, a protected sheet in words).

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { uninstallFakeHost, type FakeHelpers } from "./fakehost";
import {
  bootStyleHost,
  rejects,
  REGISTRY_KEY,
  storedNames,
  type StyleHost,
} from "./custom-styles.support";
import { STARTER_NAMES } from "../src/custom-styles";

let host: StyleHost;
let helpers: FakeHelpers;

beforeEach(async () => {
  host = await bootStyleHost();
  helpers = host.helpers;
});

afterEach(() => {
  uninstallFakeHost();
});

const FILLS: Record<string, string> = {
  "Margin %": "#F2E8DF",
  Spread: "#CCE0FF",
  Delta: "#FFE699",
};

// Saves the three styles in the order a user would: each from a cell of its own look.
async function saveThree(): Promise<void> {
  const sources = [
    ["Margin %", "Model!B2"],
    ["Spread", "Model!C2"],
    ["Delta", "Model!D2"],
  ] as const;
  for (const [name, address] of sources) {
    helpers.setFill(address, { color: FILLS[name]!, pattern: "Solid" });
    helpers.setFont(address, { bold: name === "Margin %" });
    helpers.setActiveCell(address);
    await host.smt.saveCustomStyle(name);
  }
  helpers.select("Model!A1:A3");
}

const styleAt = (address: string): string => helpers.cell(address).style;

describe("when no style is saved", () => {
  it("says what to do and where, and writes nothing", async () => {
    helpers.select("Model!A1:A3");

    expect(await rejects(() => host.smt.cycleCustomStyle())).toBe(
      "Save a style first: Tools > Custom styles.",
    );

    for (const address of ["Model!A1", "Model!A2", "Model!A3"]) {
      expect(styleAt(address), address).toBe("Normal");
    }
    expect(host.smt.undoTarget()).toBeNull();
    expect(helpers.setting(REGISTRY_KEY)).toBeNull();
  });

  it("says the same when every listed style was deleted in Excel's own Cell Styles", async () => {
    await saveThree();
    await Excel.run(async (context) => {
      for (const name of Object.keys(FILLS)) {
        context.workbook.styles.getItem(name).delete();
      }
      await context.sync();
    });

    expect(await rejects(() => host.smt.cycleCustomStyle())).toBe(
      "Save a style first: Tools > Custom styles.",
    );
  });
});

describe("one press at a time", () => {
  beforeEach(saveThree);

  it("puts the first style on a cell wearing Normal, and names its place", async () => {
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Margin % (1 of 3)");

    for (const address of ["Model!A1", "Model!A2", "Model!A3"]) {
      expect(styleAt(address), address).toBe("Margin %");
      expect(helpers.fill(address).color, address).toBe(FILLS["Margin %"]);
    }
    expect(styleAt("Model!B1")).toBe("Normal");
  });

  it("steps through the list in order, then Normal, then the first again", async () => {
    const receipts: string[] = [];
    for (let press = 0; press < 5; press += 1) {
      receipts.push(await host.smt.cycleCustomStyle());
    }

    expect(receipts).toEqual([
      "Style: Margin % (1 of 3)",
      "Style: Spread (2 of 3)",
      "Style: Delta (3 of 3)",
      "Style: Normal",
      "Style: Margin % (1 of 3)",
    ]);
    expect(styleAt("Model!A2")).toBe("Margin %");
  });

  it("leaves the selection on Normal after the last style, with Normal's own look", async () => {
    await host.smt.applyCustomStyle("Delta");
    expect(helpers.fill("Model!A1").color).toBe(FILLS.Delta);

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Normal");

    for (const address of ["Model!A1", "Model!A2", "Model!A3"]) {
      expect(styleAt(address), address).toBe("Normal");
      expect(helpers.fill(address).pattern, address).toBe("None");
    }
  });

  it("takes the first style from a style the list does not hold: a built-in or a colleague's", async () => {
    helpers.addStyle("Good", true);
    helpers.addStyle("Colleague");
    for (const worn of ["Good", "Colleague"]) {
      helpers.setStyle("Model!A1:A3", worn);
      helpers.select("Model!A1:A3");

      expect(await host.smt.cycleCustomStyle(), worn).toBe(
        "Style: Margin % (1 of 3)",
      );
    }
  });

  it("reads the style off the active cell, not the corner of the selection", async () => {
    helpers.setStyle("Model!B3", "Spread");
    helpers.select("Model!A1:B3");
    helpers.setActiveCell("Model!B3");

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Delta (3 of 3)");

    for (const address of ["Model!A1", "Model!B1", "Model!A3", "Model!B3"]) {
      expect(styleAt(address), address).toBe("Delta");
    }
  });

  it("covers every area of a ctrl-clicked selection", async () => {
    helpers.selectAreas(["Model!A1:A2", "Model!C4"]);

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Margin % (1 of 3)");

    for (const address of ["Model!A1", "Model!A2", "Model!C4"]) {
      expect(styleAt(address), address).toBe("Margin %");
    }
    expect(styleAt("Model!A3")).toBe("Normal");
  });

  it("never writes the list", async () => {
    const before = helpers.setting(REGISTRY_KEY);
    await host.smt.cycleCustomStyle();
    await host.smt.cycleCustomStyle();

    expect(helpers.setting(REGISTRY_KEY)).toBe(before);
  });

  it("skips a listed style Excel's own Cell Styles no longer has, and counts only the others", async () => {
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Spread").delete();
      await context.sync();
    });
    const before = helpers.setting(REGISTRY_KEY);

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Margin % (1 of 2)");
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Delta (2 of 2)");
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Normal");

    expect(helpers.setting(REGISTRY_KEY)).toBe(before);
    expect(storedNames(helpers)).toEqual(["Margin %", "Spread", "Delta"]);
  });

  it("reads a style worn in another case as the same style", async () => {
    helpers.setStyle("Model!A1:A3", "SPREAD");

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Delta (3 of 3)");
  });

  // Office fires a key's handler as it is pressed; each run has to read the style the one before wrote.
  it("makes two presses a moment apart two steps, not one", async () => {
    const [first, second, third] = await Promise.all([
      host.smt.cycleCustomStyle(),
      host.smt.cycleCustomStyle(),
      host.smt.cycleCustomStyle(),
    ]);

    expect([first, second, third]).toEqual([
      "Style: Margin % (1 of 3)",
      "Style: Spread (2 of 3)",
      "Style: Delta (3 of 3)",
    ]);
    expect(styleAt("Model!A1")).toBe("Delta");
  });

  it("goes on after a press that was refused", async () => {
    helpers.protectSheet("Model", []);
    const refused = host.smt.cycleCustomStyle();
    const refusedAgain = host.smt.cycleCustomStyle();
    expect(await rejects(() => refused)).toMatch(/protected/);
    expect(await rejects(() => refusedAgain)).toMatch(/protected/);

    helpers.sheet("Model").isProtected = false;
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Margin % (1 of 3)");
  });
});

describe("the starters and the user's styles in one ring", () => {
  it("walks the starters in their own order first, then the user's styles, then Normal", async () => {
    helpers.setActiveCell("Model!C2");
    await host.smt.saveCustomStyle("Margin %");
    await host.smt.addStarterStyles();
    helpers.select("Model!A1");

    const seen: string[] = [];
    for (let press = 0; press < STARTER_NAMES.length + 3; press += 1) {
      await host.smt.cycleCustomStyle();
      seen.push(styleAt("Model!A1"));
    }

    expect(seen).toEqual([
      ...STARTER_NAMES,
      "Margin %",
      "Normal",
      STARTER_NAMES[0],
    ]);
  });
});

describe("the write path is the one Apply uses", () => {
  beforeEach(saveThree);

  it("is taken back by pls,fix Undo: the old style and the cells' own look", async () => {
    helpers.setFont("Model!A1", { bold: false, color: "#AA0000" });
    helpers.setNumberFormat("Model!A1", "0.00");
    await host.smt.cycleCustomStyle();
    expect(styleAt("Model!A1")).toBe("Margin %");

    await host.smt.undoLastAction();

    expect(styleAt("Model!A1")).toBe("Normal");
    expect(helpers.font("Model!A1")).toMatchObject({
      bold: false,
      color: "#AA0000",
    });
    expect(helpers.numberFormat("Model!A1")).toBe("0.00");
    expect(helpers.fill("Model!A1").pattern).toBe("None");
  });

  it("takes the step to Normal back too: the last style and its direct format return", async () => {
    await host.smt.applyCustomStyle("Delta");
    helpers.setNumberFormat("Model!A2", "0.0x");
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Normal");
    expect(helpers.numberFormat("Model!A2")).toBe("General");

    await host.smt.undoLastAction();

    expect(styleAt("Model!A2")).toBe("Delta");
    expect(helpers.numberFormat("Model!A2")).toBe("0.0x");
    expect(helpers.fill("Model!A2").color).toBe(FILLS.Delta);
  });

  it("answers a protected sheet in words, changes nothing and spends no Undo slot", async () => {
    helpers.protectSheet("Model", []);

    expect(await rejects(() => host.smt.cycleCustomStyle())).toBe(
      "Style cycling: this sheet is protected, nothing was changed",
    );

    expect(styleAt("Model!A1")).toBe("Normal");
    expect(host.smt.undoTarget()).toBeNull();
  });

  it("refuses to read a list a newer pls,fix saved, and changes nothing", async () => {
    helpers.setSetting(REGISTRY_KEY, JSON.stringify({ v: 2, styles: [] }));

    expect(await rejects(() => host.smt.cycleCustomStyle())).toBe(
      "The styles in this workbook were saved by a newer pls,fix. Update the add-in.",
    );

    expect(styleAt("Model!A1")).toBe("Normal");
  });
});

describe("on Excel for the web", () => {
  it("steps the same way", async () => {
    uninstallFakeHost();
    host = await bootStyleHost(true);
    helpers = host.helpers;
    await saveThree();

    expect(await host.smt.cycleCustomStyle()).toBe("Style: Margin % (1 of 3)");
    expect(await host.smt.cycleCustomStyle()).toBe("Style: Spread (2 of 3)");
    expect(styleAt("Model!A1")).toBe("Spread");
  });
});
