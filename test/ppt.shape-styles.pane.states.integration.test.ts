// @vitest-environment jsdom
// The Shape styles list of the PowerPoint pane across decks and over a list it cannot simply
// trust: read again when the Tools tab opens, from a newer pls,fix, from a store that will not
// answer, corrupt, and with entries left out. The section itself: ppt.shape-styles.pane.integration.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  encodeShapeStyles,
  NEWER_SHAPE_STYLES,
  SHAPE_STYLES_KEY,
} from "../src/shape-styles-codec";
import { uninstallFakePpt } from "./fakeppt";
import {
  bootPpt,
  byId,
  dressedShape,
  GOLD,
  helpers,
  press,
  rowNames,
  rows,
  toastText,
  typeName,
} from "./ppt.shape-styles.pane.support";

afterEach(() => {
  uninstallFakePpt();
  vi.unstubAllGlobals();
});

describe("the list across decks", () => {
  it("is read again when the Tools tab opens, so another deck's styles show", async () => {
    await bootPpt();
    expect(rows()).toHaveLength(0);
    helpers
      .storage()
      .set(
        SHAPE_STYLES_KEY,
        encodeShapeStyles([{ name: "From another deck", kind: "user" }]),
      );
    await press(byId("tab-tools"));
    expect(rowNames()).toEqual(["From another deck"]);
  });
});

describe("a list the pane cannot simply trust", () => {
  it("says a list from a newer pls,fix is newer, and refuses every press", async () => {
    const raw = JSON.stringify({ v: 2, styles: [{ name: "Mine" }] });
    await bootPpt({ styles: raw });
    expect(rows()).toHaveLength(0);
    expect(byId("shape-style-empty").textContent?.trim()).toBe(
      NEWER_SHAPE_STYLES,
    );
    dressedShape();
    typeName("Mine too");
    await press(byId("save-shape-style"));
    expect(toastText()).toBe(NEWER_SHAPE_STYLES);
    await press(byId("shape-style-starters"));
    expect(toastText()).toBe(NEWER_SHAPE_STYLES);
    expect(helpers.storage().get(SHAPE_STYLES_KEY)).toBe(raw);
  });

  it("says a store that will not answer could not be read, and never reads it as empty", async () => {
    await bootPpt({ unreadable: true });
    expect(byId("shape-style-empty").textContent?.trim()).toBe(
      "Your saved shape styles could not be read.",
    );
    dressedShape();
    typeName("Mine");
    await press(byId("save-shape-style"));
    expect(toastText()).toBe(
      "Your saved shape styles could not be read, so nothing was changed.",
    );
    expect(helpers.storage().has(SHAPE_STYLES_KEY)).toBe(false);
  });

  it("starts empty from a corrupt list and says so beside the empty line", async () => {
    await bootPpt({ styles: "{broken" });
    expect(rows()).toHaveLength(0);
    expect(byId("shape-style-empty").hidden).toBe(false);
    expect(byId("shape-style-empty").textContent?.trim()).toBe(
      "Your saved shape styles could not be read, so the list starts empty.",
    );
  });

  it("keeps showing what it could read, and says how many it left out", async () => {
    await bootPpt({
      styles: JSON.stringify({
        v: 1,
        styles: [{ name: "Good", kind: "user", fill: GOLD }, { name: "" }],
      }),
    });
    expect(rowNames()).toEqual(["Good"]);
    expect(byId("shape-style-empty").hidden).toBe(false);
    expect(byId("shape-style-empty").textContent?.trim()).toBe(
      "1 style could not be read and was left out.",
    );
  });
});
