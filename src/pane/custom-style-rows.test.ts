// @vitest-environment jsdom
// The Custom styles rows as DOM: one row per style with its three buttons, the
// preview painted inline from the workbook's colors, names as text, the empty
// and unreadable lines, and a busy pane's rows born disabled. The shipped
// taskpane.html is the DOM.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it } from "vitest";
import type { StylePreview } from "../custom-styles";
import { renderStyleRows } from "./custom-style-rows";

const MARGIN: StylePreview = {
  name: "Margin %",
  kind: "user",
  fill: "#F2E8DF",
  fontColor: "#1F1D1B",
  bold: true,
};
const NOTE: StylePreview = {
  name: "pls,fix Note",
  kind: "starter",
  fill: null,
  fontColor: "#7A7A76",
  bold: false,
};
const HIGHLIGHT: StylePreview = {
  name: "pls,fix Highlight",
  kind: "starter",
  fill: "#E9EDC9",
  fontColor: null,
  bold: false,
};

const list = (): HTMLUListElement =>
  document.getElementById("custom-style-list") as HTMLUListElement;
const empty = (): HTMLElement => document.getElementById("custom-style-empty")!;
const rows = (): HTMLLIElement[] => [
  ...list().querySelectorAll<HTMLLIElement>("li"),
];
const buttonOf = (row: HTMLElement, action: string): HTMLButtonElement =>
  row.querySelector<HTMLButtonElement>(`button[data-action="${action}"]`)!;
const rgb = (hex: string): string => {
  const [r, g, b] = [1, 3, 5].map((at) => parseInt(hex.slice(at, at + 2), 16));
  return `rgb(${String(r)}, ${String(g)}, ${String(b)})`;
};

beforeEach(() => {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
});

describe("renderStyleRows", () => {
  it("shows the empty line and no list when there are no styles", () => {
    renderStyleRows({ styles: [] });

    expect(rows()).toHaveLength(0);
    expect(list().hidden).toBe(true);
    expect(empty().hidden).toBe(false);
    expect(empty().textContent).toContain("No pls,fix styles");
  });

  it("draws one row per style, in the given order, each with its three buttons", () => {
    renderStyleRows({ styles: [MARGIN, NOTE] });

    expect(list().hidden).toBe(false);
    expect(empty().hidden).toBe(true);
    expect(rows().map((row) => row.dataset.style)).toEqual([
      "Margin %",
      "pls,fix Note",
    ]);
    for (const row of rows()) {
      expect(buttonOf(row, "style-apply")).toBeTruthy();
      expect(buttonOf(row, "style-update")).toBeTruthy();
      expect(buttonOf(row, "style-delete")).toBeTruthy();
      for (const button of row.querySelectorAll("button")) {
        expect(button.dataset.style).toBe(row.dataset.style);
        expect(button.type).toBe("button");
      }
    }
  });

  it("paints the preview inline: the fill behind the text, in its color and weight", () => {
    renderStyleRows({ styles: [MARGIN, NOTE, HIGHLIGHT] });
    const sample = (row: HTMLElement) =>
      row.querySelector<HTMLElement>(".cs-sample")!;
    const [margin, note, highlight] = rows();

    expect(sample(margin!).style.backgroundColor).toBe(rgb("#F2E8DF"));
    expect(sample(margin!).style.color).toBe(rgb("#1F1D1B"));
    expect(sample(margin!).classList.contains("is-bold")).toBe(true);

    expect(sample(note!).style.backgroundColor).toBe("");
    expect(sample(note!).style.color).toBe(rgb("#7A7A76"));
    expect(sample(note!).classList.contains("is-bold")).toBe(false);

    expect(sample(highlight!).style.backgroundColor).toBe(rgb("#E9EDC9"));
    expect(sample(highlight!).style.color).toBe("");
  });

  it("keeps the preview out of the button's words: the apply button reads as the name", () => {
    renderStyleRows({ styles: [MARGIN] });
    const apply = buttonOf(rows()[0]!, "style-apply");

    expect(apply.textContent?.trim()).toBe("Margin %");
    expect(apply.querySelector(".cs-sample")?.getAttribute("aria-hidden")).toBe(
      "true",
    );
    expect(apply.title).toBe("Put Margin % on the selected cells");
  });

  it("labels the two icon buttons for a screen reader and a tooltip", () => {
    renderStyleRows({ styles: [MARGIN] });
    const [row] = rows();
    const update = buttonOf(row!, "style-update");
    const remove = buttonOf(row!, "style-delete");

    expect(update.getAttribute("aria-label")).toBe(
      "Update from the active cell",
    );
    expect(update.title).toBe("Update Margin % from the active cell");
    expect(remove.getAttribute("aria-label")).toBe("Delete");
    expect(remove.title).toBe("Delete Margin %");
    expect(update.classList.contains("icon-button")).toBe(true);
    expect(remove.classList.contains("icon-button")).toBe(true);
  });

  it("writes a name as text, never as markup", () => {
    const odd = { ...MARGIN, name: '<b onclick="x()">"Q" & co</b>' };
    renderStyleRows({ styles: [odd] });
    const row = rows()[0]!;

    expect(row.querySelector("b")).toBeNull();
    expect(row.querySelector(".cs-name")?.textContent).toBe(odd.name);
    expect(row.dataset.style).toBe(odd.name);
  });

  it("says why there is no list when the workbook's list could not be read", () => {
    renderStyleRows({ styles: [], note: "Saved by a newer pls,fix." });

    expect(list().hidden).toBe(true);
    expect(empty().hidden).toBe(false);
    expect(empty().textContent).toBe("Saved by a newer pls,fix.");
  });

  it("draws the rows disabled while the pane is busy, and back to the empty line after", () => {
    renderStyleRows({ styles: [MARGIN] }, true);
    for (const button of rows()[0]!.querySelectorAll("button")) {
      expect(button.disabled).toBe(true);
    }

    renderStyleRows({ styles: [] });
    expect(empty().textContent).toContain("No pls,fix styles");
  });
});
