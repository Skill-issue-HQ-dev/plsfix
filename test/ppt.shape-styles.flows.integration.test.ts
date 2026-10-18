// The shape styles' flows over the real adapters and the strict fake PowerPoint host: save,
// apply, update and delete, each against the user's list in the key store. Strict loads are
// on. The starters are in ppt.shape-styles.flows.starters.*, the list's states (corrupt,
// newer, refused) in ppt.shape-styles.flows.store.*.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  SHAPE_STYLES_MAX,
  starterShapeStyles,
  STARTER_SHAPE_NAMES,
  type ShapeStyle,
} from "../src/shape-styles";
import { SHAPE_STYLES_KEY } from "../src/shape-styles-codec";
import { tint } from "../src/settings";
import { uninstallFakePpt } from "./fakeppt";
import {
  bootFlows,
  callout,
  deps,
  flows,
  helpers,
  list,
  MARBLE,
  GOLD,
  names,
  palette,
  rejection,
  seed,
  shape,
  stored,
} from "./ppt.shape-styles.support";

beforeEach(bootFlows);
afterEach(uninstallFakePpt);

describe("Save style", () => {
  it("keeps the selected shape's look under the name and answers Saved style", async () => {
    callout();
    expect(await flows.saveShapeStyle(deps(), "  Margin box ")).toBe(
      "Saved style Margin box",
    );
    expect(list()).toEqual([
      {
        name: "Margin box",
        kind: "user",
        fill: "#EBEDDE",
        fillTransparency: 0,
        line: {
          visible: true,
          color: MARBLE,
          weight: 1,
          dashStyle: "Solid",
          transparency: 0,
          style: "Single",
        },
        text: { name: "Aptos", size: 14, color: "#112233", bold: true },
      },
    ]);
  });

  it("puts a new style after the ones that are there, starters first", async () => {
    seed(starterShapeStyles(palette));
    callout();
    await flows.saveShapeStyle(deps(), "Mine");
    callout();
    await flows.saveShapeStyle(deps(), "Newer");
    expect(names()).toEqual([...STARTER_SHAPE_NAMES, "Mine", "Newer"]);
  });

  it("refuses an empty name before it asks PowerPoint anything", async () => {
    helpers.selectShapes([]);
    const from = helpers.syncCount();
    expect(await rejection(() => flows.saveShapeStyle(deps(), "  "))).toBe(
      "Type a name for the style first.",
    );
    expect(helpers.syncCount()).toBe(from);
    expect(helpers.storage().has(SHAPE_STYLES_KEY)).toBe(false);
  });

  it("refuses a name that is too long", async () => {
    callout();
    expect(
      await rejection(() => flows.saveShapeStyle(deps(), "x".repeat(41))),
    ).toBe("A style name can be up to 40 characters.");
  });

  it("refuses a name already in the list, whatever its case, and keeps the list as it was", async () => {
    callout();
    await flows.saveShapeStyle(deps(), "Margin box");
    const before = helpers.storage().get(SHAPE_STYLES_KEY);
    callout();
    expect(
      await rejection(() => flows.saveShapeStyle(deps(), "MARGIN BOX")),
    ).toBe("A shape style named MARGIN BOX already exists.");
    expect(helpers.storage().get(SHAPE_STYLES_KEY)).toBe(before);
  });

  it("refuses the 31st style in words", async () => {
    seed(
      Array.from({ length: SHAPE_STYLES_MAX }, (_, at) => ({
        name: `Style ${String(at)}`,
        kind: "user" as const,
      })),
    );
    callout();
    expect(
      await rejection(() => flows.saveShapeStyle(deps(), "One more")),
    ).toBe("The list holds up to 30 shape styles. Delete one first.");
    expect(list()).toHaveLength(SHAPE_STYLES_MAX);
  });

  it("wants one selected shape, in its own words, and saves nothing", async () => {
    helpers.selectShapes([]);
    expect(await rejection(() => flows.saveShapeStyle(deps(), "Nothing"))).toBe(
      "Select one object to save its style.",
    );
    helpers.selectShapes([shape(), shape()]);
    expect(await rejection(() => flows.saveShapeStyle(deps(), "Two"))).toBe(
      "Select one object to save its style.",
    );
    expect(helpers.storage().has(SHAPE_STYLES_KEY)).toBe(false);
  });

  it("refuses a picture fill in its own words", async () => {
    helpers.selectShapes([shape({ fillImage: "<png>" })]);
    expect(await rejection(() => flows.saveShapeStyle(deps(), "Photo"))).toBe(
      "Shape styles currently support solid or no-fill objects.",
    );
  });
});

describe("a style on the selected shapes", () => {
  const callouts = (): ShapeStyle[] => starterShapeStyles(palette);

  it("puts the look on every selected shape and answers the receipt", async () => {
    seed(callouts());
    const ids = [shape({ text: "A" }), shape({ text: "B" }), shape()];
    helpers.selectShapes(ids);
    expect(await flows.applyShapeStyle(deps(), "Callout")).toBe(
      "Style Callout on 3 objects.",
    );
    for (const id of ids) {
      expect(stored(id).fillColor, id).toBe(tint(MARBLE, 0.85));
      expect(stored(id).lineColor, id).toBe(MARBLE);
      expect(stored(id).lineWeight, id).toBe(1);
      expect(stored(id).font.color, id).toBe(palette.formula);
    }
  });

  it("says object, not objects, for one", async () => {
    seed(callouts());
    helpers.selectShapes([shape()]);
    expect(await flows.applyShapeStyle(deps(), "Brand fill")).toBe(
      "Style Brand fill on 1 object.",
    );
  });

  it("counts a shape the host refuses as skipped", async () => {
    seed(callouts());
    const ids = [shape(), shape(), shape()];
    helpers.selectShapes(ids);
    helpers.refuseFormatOn(ids[1]!);
    expect(await flows.applyShapeStyle(deps(), "Brand fill")).toBe(
      "Style Brand fill on 2 objects, 1 skipped.",
    );
    expect(stored(ids[1]!).fillColor).toBeNull();
  });

  it("says in its own words that nothing took it", async () => {
    seed(callouts());
    const ids = [shape(), shape()];
    helpers.selectShapes(ids);
    for (const id of ids) helpers.refuseFormatOn(id);
    expect(
      await rejection(() => flows.applyShapeStyle(deps(), "Callout")),
    ).toBe("Style Callout does not apply to the selected objects.");
    helpers.selectShapes([ids[0]!]);
    expect(
      await rejection(() => flows.applyShapeStyle(deps(), "Callout")),
    ).toBe("Style Callout does not apply to the selected object.");
  });

  it("wants a selection", async () => {
    seed(callouts());
    helpers.selectShapes([]);
    expect(
      await rejection(() => flows.applyShapeStyle(deps(), "Callout")),
    ).toBe("Select at least one object to style.");
  });

  it("says a style that has left the list is gone, and asks PowerPoint nothing", async () => {
    seed(callouts());
    helpers.selectShapes([shape()]);
    const from = helpers.syncCount();
    expect(await rejection(() => flows.applyShapeStyle(deps(), "Nope"))).toBe(
      "The style Nope is no longer in your list.",
    );
    expect(helpers.syncCount()).toBe(from);
  });

  it("finds the style whatever the case of its name", async () => {
    seed(callouts());
    helpers.selectShapes([shape()]);
    expect(await flows.applyShapeStyle(deps(), "callout")).toBe(
      "Style Callout on 1 object.",
    );
  });
});

describe("Update from the selected shape", () => {
  it("rewrites the style from the shape and keeps its name, kind and place", async () => {
    seed([
      { name: "A", kind: "user", fill: GOLD },
      { name: "B", kind: "user", fill: GOLD, text: { bold: true } },
      { name: "C", kind: "user", fill: GOLD },
    ]);
    callout();
    expect(await flows.updateShapeStyle(deps(), "b")).toBe("Updated style B.");
    expect(names()).toEqual(["A", "B", "C"]);
    const [, updated] = list();
    expect(updated).toMatchObject({
      name: "B",
      kind: "user",
      fill: "#EBEDDE",
      text: { name: "Aptos", size: 14, color: "#112233", bold: true },
    });
  });

  it("files an updated starter under the user's own styles and says it is theirs now", async () => {
    seed(starterShapeStyles(palette));
    callout();
    expect(await flows.updateShapeStyle(deps(), "Callout")).toBe(
      "Updated Callout: it is yours now.",
    );
    const updated = list().find((style) => style.name === "Callout")!;
    expect(updated.kind).toBe("user");
    expect(updated.fill).toBe("#EBEDDE");
    expect(updated.text).toMatchObject({ bold: true, color: "#112233" });
    // The other five are still starters, in their own order, and it follows them.
    expect(names()).toEqual([
      "Brand fill",
      "Accent fill",
      "Soft fill",
      "Brand outline",
      "Highlight",
      "Callout",
    ]);
    expect(list().filter((style) => style.kind === "starter")).toHaveLength(5);
  });

  it("names the starter as the list holds it, whatever the case it was typed in", async () => {
    seed(starterShapeStyles(palette));
    callout();
    expect(await flows.updateShapeStyle(deps(), "callout")).toBe(
      "Updated Callout: it is yours now.",
    );
  });

  it("answers an ordinary update once the style is the user's: it only says it was updated", async () => {
    seed(starterShapeStyles(palette));
    callout();
    await flows.updateShapeStyle(deps(), "Callout");
    callout();
    expect(await flows.updateShapeStyle(deps(), "Callout")).toBe(
      "Updated style Callout.",
    );
    expect(list().find((style) => style.name === "Callout")!.kind).toBe("user");
  });

  it("changes nothing in the list for a starter when the update is refused", async () => {
    seed(starterShapeStyles(palette));
    const before = helpers.storage().get(SHAPE_STYLES_KEY);
    helpers.selectShapes([]);
    expect(
      await rejection(() => flows.updateShapeStyle(deps(), "Callout")),
    ).toBe("Select one object to update the style from.");
    expect(helpers.storage().get(SHAPE_STYLES_KEY)).toBe(before);
    expect(list().every((style) => style.kind === "starter")).toBe(true);
  });

  it("makes the style what the shape looks like: a shape with no text drops the text look", async () => {
    seed([{ name: "A", kind: "user", fill: GOLD, text: { color: "#FFFFFF" } }]);
    const id = shape();
    stored(id).fillColor = "#112233";
    helpers.selectShapes([id]);
    await flows.updateShapeStyle(deps(), "A");
    expect(list()[0]).not.toHaveProperty("text");
  });

  it("says a style that has left the list is gone", async () => {
    seed([{ name: "A", kind: "user" }]);
    callout();
    expect(await rejection(() => flows.updateShapeStyle(deps(), "Nope"))).toBe(
      "The style Nope is no longer in your list.",
    );
  });

  it("wants one selected shape, in its own words", async () => {
    seed([{ name: "A", kind: "user" }]);
    helpers.selectShapes([]);
    expect(await rejection(() => flows.updateShapeStyle(deps(), "A"))).toBe(
      "Select one object to update the style from.",
    );
  });
});

describe("Delete", () => {
  it("takes the style out of the list", async () => {
    seed([
      { name: "A", kind: "user" },
      { name: "B", kind: "user" },
    ]);
    expect(await flows.deleteShapeStyle(deps(), "a")).toBe("Deleted style A.");
    expect(names()).toEqual(["B"]);
  });

  it("says a style that was already gone was already gone, and writes nothing", async () => {
    seed([{ name: "A", kind: "user" }]);
    const writes = vi.spyOn(OfficeRuntime!.storage, "setItem");
    expect(await flows.deleteShapeStyle(deps(), "Nope")).toBe(
      "The style Nope was already gone.",
    );
    expect(writes).not.toHaveBeenCalled();
  });
});
