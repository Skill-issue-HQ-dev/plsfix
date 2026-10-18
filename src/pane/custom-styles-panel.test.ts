// @vitest-environment jsdom
// The Custom styles section of the Excel pane: rows drawn from what Excel holds, Enter in the name box,
// one delegated click handler, the two-press Delete that names the cells it will reset, and the list
// refreshing after every action. ../excel is mocked; the shipped taskpane.html is the DOM.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  addStarterStyles,
  applyCustomStyle,
  countStyleWearers,
  deleteCustomStyle,
  listCustomStyles,
  saveCustomStyle,
  updateCustomStyle,
} from "../excel";
import type { StylePreview } from "../custom-styles";
import { CONFIRM_MS } from "../ui/confirm";
import { isExcelReady } from "./shared";
import { renderStyleRows } from "./custom-style-rows";
import {
  addStarters,
  applyStyle,
  deleteStyle,
  installCustomStylesPanel,
  refreshCustomStyles,
  saveStyle,
  updateStyle,
} from "./custom-styles-panel";

vi.mock("../excel", () => ({
  addStarterStyles: vi.fn(),
  applyCustomStyle: vi.fn(),
  countStyleWearers: vi.fn(),
  deleteCustomStyle: vi.fn(),
  listCustomStyles: vi.fn(),
  saveCustomStyle: vi.fn(),
  updateCustomStyle: vi.fn(),
}));
vi.mock("./shared", () => ({ isExcelReady: vi.fn(() => true) }));

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

const list = (): HTMLUListElement =>
  document.getElementById("custom-style-list") as HTMLUListElement;
const empty = (): HTMLElement => document.getElementById("custom-style-empty")!;
const field = (): HTMLInputElement =>
  document.getElementById("style-name") as HTMLInputElement;
const rows = (): HTMLLIElement[] => [
  ...list().querySelectorAll<HTMLLIElement>("li"),
];
const buttonOf = (row: HTMLElement, action: string): HTMLButtonElement =>
  row.querySelector<HTMLButtonElement>(`button[data-action="${action}"]`)!;

async function settle(): Promise<void> {
  for (let i = 0; i < 6; i += 1) await Promise.resolve();
}

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(isExcelReady).mockReturnValue(true);
  vi.mocked(listCustomStyles).mockResolvedValue({ styles: [] });
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
});

describe("refreshCustomStyles", () => {
  it("draws the list Excel answers", async () => {
    vi.mocked(listCustomStyles).mockResolvedValue({ styles: [MARGIN] });
    await refreshCustomStyles();

    expect(rows().map((row) => row.dataset.style)).toEqual(["Margin %"]);
  });

  it("leaves the pane alone when Excel is not connected", async () => {
    vi.mocked(isExcelReady).mockReturnValue(false);
    await refreshCustomStyles();

    expect(listCustomStyles).not.toHaveBeenCalled();
  });

  it("says so, without throwing, when the read fails", async () => {
    vi.mocked(listCustomStyles).mockRejectedValue(new Error("boom"));

    await expect(refreshCustomStyles()).resolves.toBeUndefined();
    expect(rows()).toHaveLength(0);
    expect(empty().hidden).toBe(false);
    expect(empty().textContent).toBe(
      "Your styles could not be read from this workbook.",
    );
  });

  it("draws the rows of a busy pane disabled", async () => {
    vi.mocked(listCustomStyles).mockResolvedValue({ styles: [MARGIN] });
    (document.getElementById("style-save") as HTMLButtonElement).disabled =
      true;
    await refreshCustomStyles();

    expect(buttonOf(rows()[0]!, "style-apply").disabled).toBe(true);
  });

  it("lets the newest read win when two overlap", async () => {
    let releaseOld: (value: { styles: StylePreview[] }) => void = () =>
      undefined;
    vi.mocked(listCustomStyles)
      .mockImplementationOnce(
        () => new Promise((resolve) => (releaseOld = resolve)),
      )
      .mockResolvedValueOnce({ styles: [NOTE] });

    const older = refreshCustomStyles();
    await refreshCustomStyles();
    releaseOld({ styles: [MARGIN] });
    await older;

    expect(rows().map((row) => row.dataset.style)).toEqual(["pls,fix Note"]);
  });
});

describe("the name box and the row buttons", () => {
  it("Enter in the name box presses Save style, through the same button", () => {
    installCustomStylesPanel(vi.fn());
    const save = document.getElementById("style-save") as HTMLButtonElement;
    const clicked = vi.fn();
    save.addEventListener("click", clicked);

    field().dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );

    expect(clicked).toHaveBeenCalledTimes(1);
  });

  it("ignores every other key, and Enter while a word is still being composed", () => {
    installCustomStylesPanel(vi.fn());
    const save = document.getElementById("style-save") as HTMLButtonElement;
    const clicked = vi.fn();
    save.addEventListener("click", clicked);

    field().dispatchEvent(
      new KeyboardEvent("keydown", { key: "a", bubbles: true }),
    );
    field().dispatchEvent(
      new KeyboardEvent("keydown", {
        key: "Enter",
        isComposing: true,
        bubbles: true,
      }),
    );

    expect(clicked).not.toHaveBeenCalled();
  });

  it("hands a row button's press to run with the action and the style's name", () => {
    const run = vi.fn();
    installCustomStylesPanel(run);
    renderStyleRows({ styles: [MARGIN, NOTE] });
    const [first, second] = rows();

    buttonOf(first!, "style-apply").click();
    buttonOf(second!, "style-update").click();
    buttonOf(first!, "style-delete").click();
    // The preview inside the apply button is part of the button.
    first!.querySelector<HTMLElement>(".cs-sample")!.click();

    expect(run.mock.calls).toEqual([
      ["style-apply", "Margin %"],
      ["style-update", "pls,fix Note"],
      ["style-delete", "Margin %"],
      ["style-apply", "Margin %"],
    ]);
  });

  it("does nothing for a click that is not on a row button", () => {
    const run = vi.fn();
    installCustomStylesPanel(run);
    renderStyleRows({ styles: [MARGIN] });

    rows()[0]!.click();
    list().click();

    expect(run).not.toHaveBeenCalled();
  });

  it("installs one handler however often the rows are drawn again", () => {
    const run = vi.fn();
    installCustomStylesPanel(run);
    renderStyleRows({ styles: [MARGIN] });
    renderStyleRows({ styles: [MARGIN] });

    buttonOf(rows()[0]!, "style-apply").click();

    expect(run).toHaveBeenCalledTimes(1);
  });
});

describe("save, apply, update, starters", () => {
  beforeEach(() => {
    vi.mocked(listCustomStyles).mockResolvedValue({ styles: [MARGIN] });
  });

  it("save reads the box, saves, clears it and draws the new row", async () => {
    vi.mocked(saveCustomStyle).mockResolvedValue("Saved style Margin %");
    field().value = "  Margin %  ";

    expect(await saveStyle()).toBe("Saved style Margin %");

    expect(saveCustomStyle).toHaveBeenCalledWith("  Margin %  ");
    expect(field().value).toBe("");
    expect(rows().map((row) => row.dataset.style)).toEqual(["Margin %"]);
  });

  it("save keeps what was typed when the name is refused", async () => {
    vi.mocked(saveCustomStyle).mockRejectedValue(
      new Error("A style named Margin % already exists in this workbook."),
    );
    field().value = "Margin %";

    await expect(saveStyle()).rejects.toThrow("already exists");

    expect(field().value).toBe("Margin %");
  });

  it("apply and update call Excel with the name and draw the list again", async () => {
    vi.mocked(applyCustomStyle).mockResolvedValue("Style: Margin %");
    vi.mocked(updateCustomStyle).mockResolvedValue(
      "Updated Margin %: 2 cells wear it.",
    );

    expect(await applyStyle("Margin %")).toBe("Style: Margin %");
    expect(await updateStyle("Margin %")).toBe(
      "Updated Margin %: 2 cells wear it.",
    );

    expect(applyCustomStyle).toHaveBeenCalledWith("Margin %");
    expect(updateCustomStyle).toHaveBeenCalledWith("Margin %");
    expect(listCustomStyles).toHaveBeenCalledTimes(2);
  });

  it("draws the list again even when the action is refused (a style Excel dropped)", async () => {
    vi.mocked(applyCustomStyle).mockRejectedValue(
      new Error("The style Margin % is no longer in this workbook."),
    );

    await expect(applyStyle("Margin %")).rejects.toThrow(
      "no longer in this workbook",
    );

    expect(listCustomStyles).toHaveBeenCalledTimes(1);
  });

  it("starter styles are added, the list drawn and the receipt returned", async () => {
    vi.mocked(addStarterStyles).mockResolvedValue("Added 8 starter styles.");

    expect(await addStarters()).toBe("Added 8 starter styles.");
    expect(listCustomStyles).toHaveBeenCalledTimes(1);
  });
});

describe("delete asks twice and says what it will do", () => {
  beforeEach(async () => {
    vi.mocked(listCustomStyles).mockResolvedValue({ styles: [MARGIN, NOTE] });
    await refreshCustomStyles();
  });

  const removeOf = (index: number): HTMLButtonElement =>
    buttonOf(rows()[index]!, "style-delete");

  it("the first press counts the cells and arms the button, naming the count", async () => {
    vi.mocked(countStyleWearers).mockResolvedValue(12);

    expect(await deleteStyle("Margin %")).toBe(
      "Click again: 12 cells go back to Normal.",
    );

    expect(countStyleWearers).toHaveBeenCalledWith("Margin %");
    expect(deleteCustomStyle).not.toHaveBeenCalled();
    expect(removeOf(0).classList.contains("armed")).toBe(true);
    expect(removeOf(0).getAttribute("aria-label")).toBe(
      "Click again: 12 cells go back to Normal",
    );
    expect(removeOf(1).classList.contains("armed")).toBe(false);
  });

  it("says one cell in the singular and plain 'delete' when no one wears it", async () => {
    vi.mocked(countStyleWearers)
      .mockResolvedValueOnce(1)
      .mockResolvedValueOnce(0);

    expect(await deleteStyle("Margin %")).toBe(
      "Click again: 1 cell goes back to Normal.",
    );
    expect(await deleteStyle("pls,fix Note")).toBe("Click again to delete.");
    expect(removeOf(1).title).toBe("Click again to delete");
  });

  it("the second press deletes, draws the list again and reports", async () => {
    vi.mocked(countStyleWearers).mockResolvedValue(12);
    vi.mocked(deleteCustomStyle).mockResolvedValue(
      "Deleted style Margin %: 12 cells went back to Normal.",
    );
    await deleteStyle("Margin %");
    vi.mocked(listCustomStyles).mockResolvedValue({ styles: [NOTE] });

    expect(await deleteStyle("Margin %")).toBe(
      "Deleted style Margin %: 12 cells went back to Normal.",
    );

    expect(deleteCustomStyle).toHaveBeenCalledWith("Margin %");
    expect(countStyleWearers).toHaveBeenCalledTimes(1);
    expect(rows().map((row) => row.dataset.style)).toEqual(["pls,fix Note"]);
  });

  it("does not arm when the cells cannot be counted, and says why", async () => {
    vi.mocked(countStyleWearers).mockRejectedValue(
      new Error("Some sheets are too large to scan (Model)."),
    );

    await expect(deleteStyle("Margin %")).rejects.toThrow("too large to scan");

    expect(removeOf(0).classList.contains("armed")).toBe(false);
  });

  it("lapses on its own after the confirm window", async () => {
    vi.useFakeTimers();
    try {
      vi.mocked(countStyleWearers).mockResolvedValue(3);
      await deleteStyle("Margin %");
      expect(removeOf(0).classList.contains("armed")).toBe(true);

      vi.advanceTimersByTime(CONFIRM_MS + 1);

      expect(removeOf(0).classList.contains("armed")).toBe(false);
      expect(removeOf(0).title).toBe("Delete Margin %");
    } finally {
      vi.useRealTimers();
    }
  });

  it("drops the armed state when the rows are drawn again", async () => {
    vi.mocked(countStyleWearers).mockResolvedValue(2);
    await deleteStyle("Margin %");

    await refreshCustomStyles();
    await settle();

    expect(removeOf(0).classList.contains("armed")).toBe(false);
  });

  it("says the style is gone when its row was redrawn away while counting", async () => {
    vi.mocked(countStyleWearers).mockImplementation(async () => {
      vi.mocked(listCustomStyles).mockResolvedValue({ styles: [NOTE] });
      await refreshCustomStyles();
      return 4;
    });

    expect(await deleteStyle("Margin %")).toBe(
      "Margin % is no longer in the list.",
    );
  });
});
