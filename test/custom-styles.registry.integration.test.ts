// The registry of pls,fix styles against what Excel's Cell Styles really holds: a save the
// host refuses halfway, the starters' reserved prefix and a colleague's style of a starter's
// name, and styles deleted in Excel. Against the strict fake host.

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

describe("a save the host refuses halfway", () => {
  // The number format lands, the next property write is refused.
  async function refusedSave(name = "Margin %"): Promise<string> {
    helpers.setActiveCell("Model!B2");
    helpers.refuseStyleWriteAfter(1);
    return rejects(() => host.smt.saveCustomStyle(name));
  }

  it("lists the style it had already added, so it is not a ghost in Cell Styles", async () => {
    expect(await refusedSave()).toMatch(/argument is invalid/i);

    expect(await styleNames()).toContain("Margin %");
    expect(storedNames(helpers)).toEqual(["Margin %"]);
    expect(storedKinds(helpers)).toEqual(["user"]);
    const { styles } = await host.smt.listCustomStyles();
    expect(styles.map((style) => style.name)).toEqual(["Margin %"]);
  });
});

describe("a name in the starters' own prefix", () => {
  it("is refused for a style of the user's, and nothing is written", async () => {
    helpers.setActiveCell("Model!B2");

    expect(
      await rejects(() => host.smt.saveCustomStyle("pls,fix Margin")),
    ).toBe(
      'Style names starting with "pls,fix " are for the starter styles. Pick another name.',
    );

    expect(await styleNames()).not.toContain("pls,fix Margin");
    expect(storedNames(helpers)).toEqual([]);
  });
});

describe("Starter styles and a style of a starter's name that pls,fix did not make", () => {
  const LEFT_ALONE =
    "Left alone pls,fix Header: already in this workbook, not made by pls,fix.";

  // A colleague's style, made in Excel's own Cell Styles gallery.
  async function colleaguesHeader(): Promise<void> {
    await Excel.run(async (context) => {
      context.workbook.styles.add("pls,fix Header");
      const style = context.workbook.styles.getItem("pls,fix Header");
      style.fill.color = "#FF0000";
      style.font.bold = false;
      await context.sync();
    });
  }

  it("leaves it exactly as it is, adds the other seven and says so", async () => {
    await colleaguesHeader();
    const before = await readStyle("pls,fix Header");

    expect(await host.smt.addStarterStyles()).toBe(
      `Added 7 starter styles. ${LEFT_ALONE}`,
    );

    expect(await readStyle("pls,fix Header")).toEqual(before);
    expect(storedNames(helpers)).toEqual(
      STARTER_NAMES.filter((name) => name !== "pls,fix Header"),
    );
  });

  it("is left alone again on the next press, which restyles only the seven", async () => {
    await colleaguesHeader();
    await host.smt.addStarterStyles();
    const before = await readStyle("pls,fix Header");

    expect(await host.smt.addStarterStyles()).toBe(
      `Restyled 7 starter styles from the palette. ${LEFT_ALONE}`,
    );

    expect(await readStyle("pls,fix Header")).toEqual(before);
  });

  it("is left alone when the list calls it a user's, named as yours, and stays a user's", async () => {
    await Excel.run(async (context) => {
      context.workbook.styles.add("pls,fix Input");
      await context.sync();
    });
    helpers.setSetting(
      REGISTRY_KEY,
      JSON.stringify({
        v: 1,
        styles: [{ name: "pls,fix Input", kind: "user" }],
      }),
    );

    expect(await host.smt.addStarterStyles()).toBe(
      "Added 7 starter styles. pls,fix Input is yours now, so the starters leave it alone.",
    );

    expect(storedNames(helpers)).toContain("pls,fix Input");
    const kinds = new Map(
      storedNames(helpers).map((name, at) => [name, storedKinds(helpers)[at]]),
    );
    expect(kinds.get("pls,fix Input")).toBe("user");
    expect(kinds.get("pls,fix Title")).toBe("starter");
  });

  it("adds none and says so when all eight are someone else's", async () => {
    await Excel.run(async (context) => {
      for (const name of STARTER_NAMES) context.workbook.styles.add(name);
      await context.sync();
    });

    const receipt = await host.smt.addStarterStyles();

    expect(receipt).toBe(
      `No starter styles were added. Left alone ${STARTER_NAMES.join(", ")}: already in this workbook, not made by pls,fix.`,
    );
    expect(storedNames(helpers)).toEqual([]);
  });

  it("still restyles a starter pls,fix itself made", async () => {
    await host.smt.addStarterStyles();
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("pls,fix Header").fill.color = "#FF0000";
      await context.sync();
    });

    expect(await host.smt.addStarterStyles()).toBe(
      "Restyled 8 starter styles from the palette.",
    );
    expect((await readStyle("pls,fix Header")).fill.color).not.toBe("#FF0000");
  });
});

describe("a style deleted in Excel's own Cell Styles", () => {
  // Two user styles saved, then one of them deleted behind pls,fix's back.
  async function spreadDeletedInExcel(): Promise<void> {
    helpers.setActiveCell("Model!B2");
    await host.smt.saveCustomStyle("Margin %");
    await host.smt.saveCustomStyle("Spread");
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Spread").delete();
      await context.sync();
    });
    expect(storedNames(helpers)).toEqual(["Margin %", "Spread"]);
  }

  it("is dropped from the list the next time a style is saved", async () => {
    await spreadDeletedInExcel();

    await host.smt.saveCustomStyle("Third");

    expect(storedNames(helpers)).toEqual(["Margin %", "Third"]);
  });

  it("is dropped from the list when Starter styles runs, the starters first", async () => {
    await spreadDeletedInExcel();

    await host.smt.addStarterStyles();

    expect(storedNames(helpers)).toEqual([...STARTER_NAMES, "Margin %"]);
    expect(storedKinds(helpers).at(-1)).toBe("user");
  });

  it("is left out of what the list shows, whatever case the list spells a name in", async () => {
    helpers.setActiveCell("Model!B2");
    await host.smt.saveCustomStyle("Margin %");
    await host.smt.saveCustomStyle("Spread");
    helpers.setSetting(
      REGISTRY_KEY,
      JSON.stringify({
        v: 1,
        styles: [
          { name: "margin %", kind: "user" },
          { name: "SPREAD", kind: "user" },
          { name: "Ghost", kind: "user" },
        ],
      }),
    );
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Spread").delete();
      await context.sync();
    });

    const { styles } = await host.smt.listCustomStyles();

    expect(styles.map((style) => style.name)).toEqual(["Margin %"]);
  });
});
