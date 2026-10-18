// A starter style updated from the active cell is the user's own from then on (since 07.10.2026):
// the list files it as a user style, the receipt says so, and Starter styles from palette leaves it
// alone and names it instead of repainting it. Against the strict fake host.

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

const HEADER = "pls,fix Header";
const YOURS_NOW =
  "Updated pls,fix Header: it is yours now, Starter styles from palette leaves it alone.";
const YOURS_ALONE =
  "pls,fix Header is yours now, so the starters leave it alone.";
const NOT_MADE =
  "Left alone pls,fix Note: already in this workbook, not made by pls,fix.";
const LOUD = "#CCE0FF";

let host: StyleHost;
let helpers: FakeHelpers;

beforeEach(async () => {
  host = await bootStyleHost();
  helpers = host.helpers;
  await host.smt.addStarterStyles();
});

afterEach(() => {
  uninstallFakeHost();
});

// A look that says something else in every field a style carries, in C3, the active cell.
function loudLook(): void {
  helpers.setFill("Model!C3", { color: LOUD, pattern: "Solid" });
  helpers.setFont("Model!C3", {
    name: "Courier New",
    size: 18,
    bold: true,
    italic: true,
    color: "#AA0000",
  });
  helpers.setNumberFormat("Model!C3", "0.00");
  helpers.setActiveCell("Model!C3");
}

function kinds(): Record<string, string> {
  return Object.fromEntries(
    storedNames(helpers).map((name, at) => [name, storedKinds(helpers)[at]!]),
  );
}

describe("updating a starter from the active cell", () => {
  it("files it as the user's own, in the receipt and in the workbook's list", async () => {
    loudLook();

    expect(await host.smt.updateCustomStyle(HEADER)).toBe(YOURS_NOW);

    expect(kinds()[HEADER]).toBe("user");
    for (const name of STARTER_NAMES.filter((other) => other !== HEADER)) {
      expect(kinds()[name], name).toBe("starter");
    }
  });

  it("still rewrites the style from the cell, and every cell wearing it follows", async () => {
    helpers.select("Model!A1:B1");
    await host.smt.applyCustomStyle(HEADER);
    loudLook();

    await host.smt.updateCustomStyle(HEADER);

    expect(await readStyle(HEADER)).toMatchObject({
      numberFormat: "0.00",
      fill: { color: LOUD },
      font: { name: "Courier New", size: 18, bold: true, color: "#AA0000" },
    });
    for (const address of ["Model!A1", "Model!B1"]) {
      expect(helpers.cell(address).style, address).toBe(HEADER);
      expect(helpers.fill(address).color, address).toBe(LOUD);
    }
  });

  it("moves it out of the starters into the user's group, ahead of the styles the user saved", async () => {
    helpers.setActiveCell("Model!B2");
    await host.smt.saveCustomStyle("Spread");
    loudLook();

    await host.smt.updateCustomStyle(HEADER);

    expect(storedNames(helpers)).toEqual([
      ...STARTER_NAMES.filter((name) => name !== HEADER),
      HEADER,
      "Spread",
    ]);
  });

  it("is a plain update the second time: the user's style, its wearers counted", async () => {
    helpers.select("Model!A1:A2");
    await host.smt.applyCustomStyle(HEADER);
    loudLook();
    await host.smt.updateCustomStyle(HEADER);

    expect(await host.smt.updateCustomStyle(HEADER)).toBe(
      "Updated pls,fix Header: 2 cells wear it.",
    );
    expect(kinds()[HEADER]).toBe("user");
  });

  it("writes the list before the style, so a host that stops halfway never leaves a starter the next press would repaint", async () => {
    loudLook();
    helpers.refuseStyleWriteAfter(1);

    expect(await rejects(() => host.smt.updateCustomStyle(HEADER))).toMatch(
      /argument is invalid/i,
    );

    expect(kinds()[HEADER]).toBe("user");
  });

  it("under a list a newer pls,fix wrote, updates the style and writes no list", async () => {
    const newer = JSON.stringify({ v: 2, styles: [] });
    helpers.setSetting(REGISTRY_KEY, newer);
    loudLook();

    expect(await host.smt.updateCustomStyle(HEADER)).toBe(
      "Updated pls,fix Header: no cell wears it yet.",
    );

    expect(helpers.setting(REGISTRY_KEY)).toBe(newer);
    expect((await readStyle(HEADER)).fill.color).toBe(LOUD);
  });

  it("leaves the list alone for a style it does not list", async () => {
    await Excel.run(async (context) => {
      context.workbook.styles.add("Colleague");
      await context.sync();
    });
    const before = helpers.setting(REGISTRY_KEY);
    loudLook();

    expect(await host.smt.updateCustomStyle("Colleague")).toBe(
      "Updated Colleague: no cell wears it yet.",
    );

    expect(helpers.setting(REGISTRY_KEY)).toBe(before);
  });
});

describe("Starter styles from palette after a starter became the user's", () => {
  beforeEach(async () => {
    loudLook();
    await host.smt.updateCustomStyle(HEADER);
  });

  it("leaves it exactly as it is, restyles the other seven and names it", async () => {
    const before = await readStyle(HEADER);

    expect(await host.smt.addStarterStyles()).toBe(
      `Restyled 7 starter styles from the palette. ${YOURS_ALONE}`,
    );

    expect(await readStyle(HEADER)).toEqual(before);
    expect(kinds()[HEADER]).toBe("user");
  });

  it("names it as the user's, never as a style pls,fix did not make", async () => {
    const receipt = await host.smt.addStarterStyles();

    expect(receipt).not.toContain("not made by pls,fix");
    expect(receipt).not.toContain("Left alone");
  });

  it("names two of them together, in the plural, in the starters' order", async () => {
    await host.smt.updateCustomStyle("pls,fix Note");

    expect(await host.smt.addStarterStyles()).toBe(
      "Restyled 6 starter styles from the palette. pls,fix Header and pls,fix Note are yours now, so the starters leave them alone.",
    );
  });

  it("names a style the list does not hold on its own, after the ones that are yours", async () => {
    const listed = JSON.parse(helpers.setting(REGISTRY_KEY)!) as {
      v: number;
      styles: { name: string; kind: string }[];
    };
    helpers.setSetting(
      REGISTRY_KEY,
      JSON.stringify({
        ...listed,
        styles: listed.styles.filter((style) => style.name !== "pls,fix Note"),
      }),
    );

    expect(await host.smt.addStarterStyles()).toBe(
      `Restyled 6 starter styles from the palette. ${YOURS_ALONE} ${NOT_MADE}`,
    );

    expect(kinds()[HEADER]).toBe("user");
    expect(kinds()["pls,fix Note"]).toBeUndefined();
  });

  it("does not recreate a starter under that name, press after press", async () => {
    await host.smt.addStarterStyles();
    await host.smt.addStarterStyles();

    expect((await styleNames()).filter((name) => name === HEADER)).toHaveLength(
      1,
    );
    expect(storedNames(helpers).filter((name) => name === HEADER)).toHaveLength(
      1,
    );
    expect(kinds()[HEADER]).toBe("user");
    expect((await readStyle(HEADER)).fill.color).toBe(LOUD);
  });

  it("follows a palette change in the other seven and keeps the updated one's look", async () => {
    host.settings.setActiveSettings({
      ...host.settings.getActiveSettings(),
      primary: "#AA0000",
      brand: ["#AA0000", host.settings.getActiveSettings().accent],
    });

    await host.smt.addStarterStyles();

    expect((await readStyle("pls,fix Title")).fill.color).toBe("#AA0000");
    expect((await readStyle(HEADER)).fill.color).toBe(LOUD);
  });

  it("frees the name when the user deletes the style: the next press adds a fresh starter", async () => {
    await host.smt.deleteCustomStyle(HEADER);
    expect(await styleNames()).not.toContain(HEADER);

    expect(await host.smt.addStarterStyles()).toBe(
      "Added 1 and restyled 7 starter styles from the palette.",
    );

    expect(kinds()[HEADER]).toBe("starter");
    expect(storedNames(helpers)).toEqual([...STARTER_NAMES]);
  });
});

describe("a starter changed in Excel's own Cell Styles, not through Update", () => {
  it("is still the starter's: the next press restyles it, and says nothing is left alone", async () => {
    await Excel.run(async (context) => {
      context.workbook.styles.getItem(HEADER).fill.color = LOUD;
      await context.sync();
    });

    expect(await host.smt.addStarterStyles()).toBe(
      "Restyled 8 starter styles from the palette.",
    );

    expect((await readStyle(HEADER)).fill.color).not.toBe(LOUD);
    expect(kinds()[HEADER]).toBe("starter");
  });
});
