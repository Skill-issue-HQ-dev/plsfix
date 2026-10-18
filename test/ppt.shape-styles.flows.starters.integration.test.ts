// The six starter shape styles from the palette, over the real adapters and the
// strict fake PowerPoint host: created when missing, restyled when there, a
// user's own style under a starter's name left alone and named, the cap of 30
// respected. The other flows are in ppt.shape-styles.flows.integration.test.ts.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  SHAPE_STYLES_MAX,
  STARTER_SHAPE_NAMES,
  starterShapeStyles,
  type ShapeStyle,
} from "../src/shape-styles";
import { uninstallFakePpt } from "./fakeppt";
import {
  bootFlows,
  callout,
  deps,
  flows,
  GOLD,
  GREEN,
  helpers,
  list,
  MARBLE,
  names,
  palette,
  seed,
  setPalette,
  settingsOf,
  shape,
  SLATE,
  stored,
} from "./ppt.shape-styles.support";

beforeEach(bootFlows);
afterEach(uninstallFakePpt);

describe("Starter styles from palette", () => {
  it("adds the six from the palette the pane shows", async () => {
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 6 starter styles.",
    );
    expect(names()).toEqual([...STARTER_SHAPE_NAMES]);
    expect(list()).toEqual(starterShapeStyles(palette));
  });

  it("restyles the six from a palette that changed, and says so", async () => {
    await flows.addStarterShapeStyles(deps());
    setPalette(settingsOf([SLATE, GREEN]));
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Restyled 6 starter styles from the palette.",
    );
    expect(list()).toEqual(starterShapeStyles(palette));
  });

  it("adds what is missing and restyles what is there in one receipt", async () => {
    seed(
      starterShapeStyles(palette).filter(
        (style) => style.name === "Callout" || style.name === "Highlight",
      ),
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 4 and restyled 2 starter styles from the palette.",
    );
    expect(names()).toEqual([...STARTER_SHAPE_NAMES]);
  });

  it("keeps the user's styles and puts them after the starters", async () => {
    seed([{ name: "Mine", kind: "user", fill: GOLD }]);
    await flows.addStarterShapeStyles(deps());
    expect(names()).toEqual([...STARTER_SHAPE_NAMES, "Mine"]);
  });

  it("leaves a style the user made under a starter's name as it is, and names it", async () => {
    const mine: ShapeStyle = { name: "Callout", kind: "user", fill: SLATE };
    seed([mine]);
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 5 starter styles. Callout is yours now, so the starters leave it alone.",
    );
    expect(list().find((style) => style.name === "Callout")).toEqual(mine);
  });

  it("names every style of the user's under a starter's name, in one sentence", async () => {
    seed(
      ["Brand fill", "Callout", "Highlight"].map((name) => ({
        name,
        kind: "user" as const,
        fill: SLATE,
      })),
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 3 starter styles. Brand fill, Callout and Highlight are yours now, so the starters leave them alone.",
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Restyled 3 starter styles from the palette. Brand fill, Callout and Highlight are yours now, so the starters leave them alone.",
    );
    seed(
      ["Callout", "Highlight"].map((name) => ({
        name,
        kind: "user" as const,
        fill: SLATE,
      })),
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 4 starter styles. Callout and Highlight are yours now, so the starters leave them alone.",
    );
  });

  it("adds what fits under the cap and names what did not", async () => {
    seed(
      Array.from({ length: 27 }, (_, at) => ({
        name: `Mine ${String(at)}`,
        kind: "user" as const,
      })),
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 3 starter styles. Left out Brand outline, Callout, Highlight: the list holds up to 30 styles.",
    );
    expect(list()).toHaveLength(SHAPE_STYLES_MAX);
  });

  it("answers in words when nothing was added", async () => {
    seed(
      Array.from({ length: SHAPE_STYLES_MAX }, (_, at) => ({
        name: `Mine ${String(at)}`,
        kind: "user" as const,
      })),
    );
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "No starter styles were added. Left out Brand fill, Accent fill, Soft fill, Brand outline, Callout, Highlight: the list holds up to 30 styles.",
    );
  });
});

// A starter the user updated is theirs (since 07.10.2026): the next press must neither
// repaint it nor make a starter under its name. What a press writes is read off the host's
// key store, and what Apply then paints off the host's shape.
describe("a starter the user updated", () => {
  async function takeOver(name: string): Promise<ShapeStyle> {
    callout();
    await flows.updateShapeStyle(deps(), name);
    return list().find((style) => style.name === name)!;
  }

  it("is left alone by the next press, which says so and restyles the other five", async () => {
    await flows.addStarterShapeStyles(deps());
    const mine = await takeOver("Callout");
    expect(mine.kind).toBe("user");
    setPalette(settingsOf([SLATE, GREEN]));
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Restyled 5 starter styles from the palette. Callout is yours now, so the starters leave it alone.",
    );
    // What the host's key store holds after the press.
    expect(list().find((style) => style.name === "Callout")).toEqual(mine);
    expect(names().filter((name) => name === "Callout")).toHaveLength(1);
    expect(list().filter((style) => style.kind === "starter")).toEqual(
      starterShapeStyles(palette).filter((style) => style.name !== "Callout"),
    );
  });

  it("is what Apply paints afterwards, not the palette's callout", async () => {
    await flows.addStarterShapeStyles(deps());
    await takeOver("Callout");
    await flows.addStarterShapeStyles(deps());
    const id = shape({ text: "Plain" });
    helpers.selectShapes([id]);
    expect(await flows.applyShapeStyle(deps(), "Callout")).toBe(
      "Style Callout on 1 object.",
    );
    expect(stored(id).fillColor).toBe("#EBEDDE");
    expect(stored(id).lineColor).toBe(MARBLE);
    expect(stored(id).font.color).toBe("#112233");
  });

  it("makes the other five and no starter under its name", async () => {
    seed(
      starterShapeStyles(palette).filter((style) => style.name === "Callout"),
    );
    await takeOver("Callout");
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 5 starter styles. Callout is yours now, so the starters leave it alone.",
    );
    expect(names()).toEqual([
      "Brand fill",
      "Accent fill",
      "Soft fill",
      "Brand outline",
      "Highlight",
      "Callout",
    ]);
    expect(list().find((style) => style.name === "Callout")!.kind).toBe("user");
  });

  it("is named with the other starters the user took over, in the starters' order", async () => {
    await flows.addStarterShapeStyles(deps());
    await takeOver("Highlight");
    await takeOver("Callout");
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Restyled 4 starter styles from the palette. Callout and Highlight are yours now, so the starters leave them alone.",
    );
  });

  it("comes back as a starter once the user deleted the style they took over", async () => {
    await flows.addStarterShapeStyles(deps());
    await takeOver("Callout");
    await flows.deleteShapeStyle(deps(), "Callout");
    expect(names()).not.toContain("Callout");
    expect(await flows.addStarterShapeStyles(deps())).toBe(
      "Added 1 and restyled 5 starter styles from the palette.",
    );
    expect(list().find((style) => style.name === "Callout")).toEqual(
      starterShapeStyles(palette).find((style) => style.name === "Callout"),
    );
  });
});
