// @vitest-environment jsdom
// The Shape styles section of the PowerPoint pane, driven the way a person drives it: the real
// pptpane.html and src/ppt/main.ts over the strict fake host and its key store. Save style
// (button and Enter), a row's apply, Update, Delete (two presses) and Starter styles. The list
// across decks and an untrustworthy list: ppt.shape-styles.pane.states.integration.test.ts.

import { afterEach, describe, expect, it, vi } from "vitest";
import { STYLE_NAME_MAX } from "../src/custom-styles";
import { DEFAULT_SETTINGS, tint } from "../src/settings";
import {
  SHAPE_STYLES_MAX,
  STARTER_SHAPE_NAMES,
  starterShapeStyles,
} from "../src/shape-styles";
import { encodeShapeStyles, SHAPE_STYLES_KEY } from "../src/shape-styles-codec";
import { uninstallFakePpt } from "./fakeppt";
import { settlePpt } from "./ppt-ready";
import {
  BRAND,
  bootPpt,
  byId,
  deck,
  dressedShape,
  GOLD,
  helpers,
  MARBLE,
  plainShapes,
  press,
  rowButton,
  rowNames,
  rows,
  storedStyles,
  toastText,
  typeName,
} from "./ppt.shape-styles.pane.support";

afterEach(() => {
  uninstallFakePpt();
  vi.unstubAllGlobals();
});

describe("the section as the pane ships it", () => {
  it("sits under Smart Painter with a short placeholder and the name limit", async () => {
    await bootPpt();
    const field = byId<HTMLInputElement>("shape-style-name");
    expect(field.maxLength).toBe(STYLE_NAME_MAX);
    // Short enough not to clip at 320 px beside the Save style button.
    expect(field.placeholder).toBe("Style name");
    expect(field.placeholder.length).toBeLessThan(16);
    const section = byId("shape-styles-heading").closest("section")!;
    expect(section.previousElementSibling).toBe(
      byId("smart-painter-heading").closest("section"),
    );
    expect(byId("save-shape-style").textContent?.trim()).toBe("Save style");
    expect(byId("shape-style-starters").textContent?.trim()).toBe(
      "Starter styles from palette",
    );
  });

  it("says there are none when the user has saved none", async () => {
    await bootPpt();
    expect(rows()).toHaveLength(0);
    expect(byId("shape-style-list").hidden).toBe(true);
    expect(byId("shape-style-empty").hidden).toBe(false);
    expect(byId("shape-style-empty").textContent?.trim()).toBe(
      "No shape styles saved yet.",
    );
  });

  it("draws the saved list at boot, each row with its own preview", async () => {
    const list = starterShapeStyles({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: BRAND,
    });
    await bootPpt({ styles: encodeShapeStyles(list) });
    expect(rowNames()).toEqual([...STARTER_SHAPE_NAMES]);
    const sample = rows()[0]!.querySelector<HTMLElement>(".cs-sample")!;
    expect(sample.style.backgroundColor).toBe("rgb(178, 126, 84)");
  });
});

describe("Save style", () => {
  it("saves the selected shape under the typed name when Enter is pressed", async () => {
    await bootPpt();
    dressedShape();
    typeName("Margin box");
    byId("shape-style-name").dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    await settlePpt();
    expect(toastText()).toBe("Saved style Margin box");
    expect(rowNames()).toEqual(["Margin box"]);
    expect(byId<HTMLInputElement>("shape-style-name").value).toBe("");
    expect(storedStyles()[0]).toMatchObject({
      name: "Margin box",
      kind: "user",
      fill: "#EBEDDE",
      line: { visible: true, color: MARBLE, weight: 2 },
      text: { name: "Aptos", size: 14, color: "#112233", bold: true },
    });
  });

  it("saves from the button too, and keeps the name in the box when it refuses", async () => {
    await bootPpt();
    helpers.selectShapes([]);
    typeName("Nothing selected");
    await press(byId("save-shape-style"));
    expect(toastText()).toBe("Select one object to save its style.");
    expect(byId<HTMLInputElement>("shape-style-name").value).toBe(
      "Nothing selected",
    );
    expect(rows()).toHaveLength(0);
  });

  it("asks for a name first", async () => {
    await bootPpt();
    dressedShape();
    await press(byId("save-shape-style"));
    expect(toastText()).toBe("Type a name for the style first.");
  });

  it("refuses a name that is in the list, whatever its case", async () => {
    await bootPpt();
    dressedShape();
    typeName("Margin box");
    await press(byId("save-shape-style"));
    dressedShape();
    typeName("MARGIN BOX");
    await press(byId("save-shape-style"));
    expect(toastText()).toBe("A shape style named MARGIN BOX already exists.");
    expect(rowNames()).toEqual(["Margin box"]);
  });

  it("refuses the 31st in words and keeps the 30", async () => {
    const thirty = Array.from({ length: SHAPE_STYLES_MAX }, (_, at) => ({
      name: `Style ${String(at)}`,
      kind: "user" as const,
    }));
    await bootPpt({ styles: encodeShapeStyles(thirty) });
    dressedShape();
    typeName("One more");
    await press(byId("save-shape-style"));
    expect(toastText()).toBe(
      "The list holds up to 30 shape styles. Delete one first.",
    );
    expect(rows()).toHaveLength(SHAPE_STYLES_MAX);
  });
});

describe("a row", () => {
  it("applies the style to every selected shape and says so", async () => {
    const list = starterShapeStyles({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: BRAND,
    });
    await bootPpt({ styles: encodeShapeStyles(list) });
    const ids = plainShapes(3);
    await press(rowButton("Callout", "shape-style-apply"));
    expect(toastText()).toBe("Style Callout on 3 objects.");
    for (const id of ids) {
      const held = deck.findShape(id).shape;
      expect(held.fillColor, id).toBe(tint(MARBLE, 0.85));
      expect(held.lineColor, id).toBe(MARBLE);
      expect(held.lineWeight, id).toBe(1);
      expect(held.font.color, id).toBe(DEFAULT_SETTINGS.formula);
    }
  });

  it("rewrites the style from the selected shape with the Update icon", async () => {
    await bootPpt({
      styles: encodeShapeStyles([{ name: "Mine", kind: "user", fill: GOLD }]),
    });
    dressedShape();
    await press(rowButton("Mine", "shape-style-update"));
    expect(toastText()).toBe("Updated style Mine.");
    expect(storedStyles()[0]).toMatchObject({
      name: "Mine",
      fill: "#EBEDDE",
      text: { color: "#112233" },
    });
    // The row is drawn again from what the list holds now.
    const sample = rows()[0]!.querySelector<HTMLElement>(".cs-sample")!;
    expect(sample.style.backgroundColor).toBe("rgb(235, 237, 222)");
  });

  it("deletes the style on the second press of the Delete icon, the first only arms it", async () => {
    await bootPpt({
      styles: encodeShapeStyles([
        { name: "A", kind: "user" },
        { name: "B", kind: "user" },
      ]),
    });
    const remove = rowButton("A", "shape-style-delete");
    await press(remove);
    expect(remove.classList.contains("armed")).toBe(true);
    expect(remove.title).toBe("Click again to confirm");
    expect(rowNames()).toEqual(["A", "B"]);
    expect(storedStyles()).toHaveLength(2);

    await press(remove);
    expect(toastText()).toBe("Deleted style A.");
    expect(rowNames()).toEqual(["B"]);
    expect(storedStyles().map((style) => style.name)).toEqual(["B"]);
  });

  it("hands every button back after an action, and holds them while one runs", async () => {
    await bootPpt({
      styles: encodeShapeStyles([{ name: "A", kind: "user", fill: GOLD }]),
    });
    plainShapes(2);
    rowButton("A", "shape-style-apply").click();
    // The busy latch went on at the press: the row's buttons cannot be pressed twice.
    for (const button of document.querySelectorAll<HTMLButtonElement>(
      "#shape-style-list button",
    )) {
      expect(button.disabled).toBe(true);
    }
    await settlePpt();
    for (const button of document.querySelectorAll<HTMLButtonElement>(
      "#shape-style-list button",
    )) {
      expect(button.disabled).toBe(false);
    }
    expect(byId<HTMLButtonElement>("save-shape-style").disabled).toBe(false);
  });

  it("says a style another deck's pane deleted is gone, and draws the list again", async () => {
    await bootPpt({
      styles: encodeShapeStyles([
        { name: "A", kind: "user", fill: GOLD },
        { name: "B", kind: "user", fill: GOLD },
      ]),
    });
    plainShapes(1);
    helpers
      .storage()
      .set(
        SHAPE_STYLES_KEY,
        encodeShapeStyles([{ name: "B", kind: "user", fill: GOLD }]),
      );
    await press(rowButton("A", "shape-style-apply"));
    expect(toastText()).toBe("The style A is no longer in your list.");
    expect(rowNames()).toEqual(["B"]);
  });
});

describe("Starter styles from palette", () => {
  it("adds the six in the palette the pane shows and draws them", async () => {
    await bootPpt();
    await press(byId("shape-style-starters"));
    expect(toastText()).toBe("Added 6 starter styles.");
    expect(rowNames()).toEqual([...STARTER_SHAPE_NAMES]);
    expect(storedStyles()[0]).toMatchObject({
      name: "Brand fill",
      kind: "starter",
      fill: GOLD,
    });
    expect(storedStyles()[1]).toMatchObject({
      name: "Accent fill",
      fill: MARBLE,
    });
  });

  it("restyles the six when pressed again", async () => {
    await bootPpt();
    await press(byId("shape-style-starters"));
    await press(byId("shape-style-starters"));
    expect(toastText()).toBe("Restyled 6 starter styles from the palette.");
    expect(rows()).toHaveLength(6);
  });

  it("restyles from the palette as it is now, after the editor changed it", async () => {
    await bootPpt();
    await press(byId("shape-style-starters"));
    document.getElementById("tab-settings")!.click();
    const picker = document.querySelector<HTMLInputElement>(
      '#ppt-brand-editor input[type="color"]',
    )!;
    picker.value = "#112233";
    picker.dispatchEvent(new Event("input", { bubbles: true }));
    picker.dispatchEvent(new Event("change", { bubbles: true }));
    await settlePpt();
    await press(byId("shape-style-starters"));
    expect(toastText()).toBe("Restyled 6 starter styles from the palette.");
    expect(storedStyles()[0]!.fill).toBe("#112233");
  });

  // Since 07.10.2026 a starter the user updates is theirs from then on.
  it("takes a starter's Update icon as the user's own and leaves it alone after", async () => {
    const list = starterShapeStyles({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: BRAND,
    });
    await bootPpt({ styles: encodeShapeStyles(list) });
    dressedShape();
    await press(rowButton("Callout", "shape-style-update"));
    expect(toastText()).toBe("Updated Callout: it is yours now.");
    expect(rowNames()).toEqual([
      "Brand fill",
      "Accent fill",
      "Soft fill",
      "Brand outline",
      "Highlight",
      "Callout",
    ]);
    const rowOf = (name: string): HTMLElement =>
      rows().find((row) => row.dataset.style === name)!;
    expect(rowOf("Callout").dataset.kind).toBe("user");
    expect(rowOf("Highlight").dataset.kind).toBe("starter");

    await press(byId("shape-style-starters"));
    expect(toastText()).toBe(
      "Restyled 5 starter styles from the palette. Callout is yours now, so the starters leave it alone.",
    );
    expect(
      storedStyles().find((style) => style.name === "Callout"),
    ).toMatchObject({
      kind: "user",
      fill: "#EBEDDE",
      text: { color: "#112233" },
    });
    expect(rowNames().filter((name) => name === "Callout")).toHaveLength(1);
  });
});
