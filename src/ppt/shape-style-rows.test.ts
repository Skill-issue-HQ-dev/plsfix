// @vitest-environment jsdom
// The Shape styles list as DOM: one row per style (an apply button with a preview
// of its fill, outline and text, an Update icon, a Delete icon) or the line that
// says why there is none. No Office.js and no pane state. A style's name only ever
// travels as text or as an attribute value.

import { beforeEach, describe, expect, it } from "vitest";
import type { ShapeStylePreview } from "../shape-styles";
import { drawShapeStyleRows, SHAPE_STYLE_ACTIONS } from "./shape-style-rows";

let list: HTMLUListElement;
let note: HTMLElement;

beforeEach(() => {
  list = document.createElement("ul");
  list.hidden = true;
  note = document.createElement("p");
  document.body.replaceChildren(list, note);
});

const preview = (over: Partial<ShapeStylePreview> = {}): ShapeStylePreview => ({
  name: "Callout",
  kind: "starter",
  fill: "#DFDEDE",
  outline: { color: "#282623", weight: 1, dashed: false },
  textColor: "#1F1D1B",
  bold: false,
  ...over,
});

const rows = (): HTMLElement[] => [
  ...list.querySelectorAll<HTMLElement>("li.cs-row"),
];
const sampleOf = (row: HTMLElement): HTMLElement =>
  row.querySelector<HTMLElement>(".cs-sample")!;

describe("drawShapeStyleRows", () => {
  it("draws one row per style, in list order, with its name and kind", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview(), preview({ name: "Mine", kind: "user" })],
    });
    expect(rows().map((row) => row.dataset.style)).toEqual(["Callout", "Mine"]);
    expect(rows().map((row) => row.dataset.kind)).toEqual(["starter", "user"]);
    expect(list.hidden).toBe(false);
    expect(note.hidden).toBe(true);
  });

  it("gives each row its apply button, Update icon and Delete icon", () => {
    drawShapeStyleRows(list, note, { styles: [preview()] });
    const [apply, update, remove] = [
      ...rows()[0]!.querySelectorAll<HTMLButtonElement>("button"),
    ];
    expect(apply!.dataset.action).toBe(SHAPE_STYLE_ACTIONS.apply);
    expect(apply!.querySelector(".cs-name")!.textContent).toBe("Callout");
    expect(apply!.title).toBe("Put Callout on the selected objects");
    expect(update!.dataset.action).toBe(SHAPE_STYLE_ACTIONS.update);
    expect(update!.getAttribute("aria-label")).toBe(
      "Update from the selected shape",
    );
    expect(update!.title).toBe("Update Callout from the selected shape");
    expect(remove!.dataset.action).toBe(SHAPE_STYLE_ACTIONS.remove);
    expect(remove!.getAttribute("aria-label")).toBe("Delete");
    for (const button of [apply!, update!, remove!]) {
      expect(button.dataset.style).toBe("Callout");
      expect(button.type).toBe("button");
    }
    expect(update!.classList.contains("icon-button")).toBe(true);
    expect(remove!.classList.contains("icon-button")).toBe(true);
  });

  it("paints the style's own fill, outline and text color on the preview", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview({ bold: true })],
    });
    const sample = sampleOf(rows()[0]!);
    expect(sample.style.backgroundColor).toBe("rgb(223, 222, 222)");
    expect(sample.style.borderColor).toBe("rgb(40, 38, 35)");
    expect(sample.style.borderStyle).toBe("solid");
    expect(sample.style.color).toBe("rgb(31, 29, 27)");
    expect(sample.classList.contains("is-bold")).toBe(true);
    expect(sample.getAttribute("aria-hidden")).toBe("true");
  });

  it("draws a dashed outline dashed and thicker for a heavier weight, never past three pixels", () => {
    drawShapeStyleRows(list, note, {
      styles: [
        preview({ outline: { color: "#112233", weight: 2, dashed: true } }),
        preview({
          name: "Heavy",
          outline: { color: "#112233", weight: 20, dashed: false },
        }),
        preview({
          name: "Hair",
          outline: { color: "#112233", weight: 0.25, dashed: false },
        }),
      ],
    });
    const [dashed, heavy, hair] = rows().map((row) => sampleOf(row));
    expect(dashed!.style.borderStyle).toBe("dashed");
    expect(dashed!.style.borderWidth).toBe("2px");
    expect(heavy!.style.borderWidth).toBe("3px");
    expect(hair!.style.borderWidth).toBe("1px");
  });

  it("paints nothing where the style has no fill, no outline or no text color", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview({ fill: null, outline: null, textColor: null })],
    });
    const sample = sampleOf(rows()[0]!);
    expect(sample.style.backgroundColor).toBe("");
    expect(sample.style.color).toBe("");
    // No outline: the preview's own hairline is drawn clear, so it cannot pass for one.
    expect(sample.style.borderColor).toBe("transparent");
  });

  it("keeps the preview's hairline for an outline whose color is the theme's", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview({ outline: { color: null, weight: 1, dashed: false } })],
    });
    expect(sampleOf(rows()[0]!).style.borderColor).toBe("");
  });

  it("disables the buttons of rows born while an action runs", () => {
    drawShapeStyleRows(list, note, { styles: [preview()] }, true);
    for (const button of list.querySelectorAll("button")) {
      expect(button.disabled).toBe(true);
    }
    drawShapeStyleRows(list, note, { styles: [preview()] }, false);
    for (const button of list.querySelectorAll("button")) {
      expect(button.disabled).toBe(false);
    }
  });

  it("replaces the rows it drew before, armed state and all", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview(), preview({ name: "B" })],
    });
    drawShapeStyleRows(list, note, { styles: [preview({ name: "C" })] });
    expect(rows().map((row) => row.dataset.style)).toEqual(["C"]);
  });

  it("says there are none, in the note, and hides the list", () => {
    drawShapeStyleRows(list, note, { styles: [] });
    expect(list.hidden).toBe(true);
    expect(note.hidden).toBe(false);
    expect(note.textContent).toBe("No shape styles saved yet.");
  });

  it("says why the list could not be read in place of the empty line", () => {
    drawShapeStyleRows(list, note, {
      styles: [],
      note: "The saved shape styles are from a newer pls,fix. Update the add-in.",
    });
    expect(note.textContent).toBe(
      "The saved shape styles are from a newer pls,fix. Update the add-in.",
    );
    expect(note.hidden).toBe(false);
  });

  it("keeps showing a note beside a list that was read with entries left out", () => {
    drawShapeStyleRows(list, note, {
      styles: [preview()],
      note: "1 style could not be read and was left out.",
    });
    expect(list.hidden).toBe(false);
    expect(note.hidden).toBe(false);
    expect(note.textContent).toBe(
      "1 style could not be read and was left out.",
    );
  });

  it("carries a name with markup, quotes or a line break as text, never as HTML", () => {
    const name = `<img src=x onerror=alert(1)> "q" & 'a'`;
    drawShapeStyleRows(list, note, { styles: [preview({ name })] });
    expect(list.querySelector("img")).toBeNull();
    expect(rows()[0]!.dataset.style).toBe(name);
    expect(rows()[0]!.querySelector(".cs-name")!.textContent).toBe(name);
  });
});
