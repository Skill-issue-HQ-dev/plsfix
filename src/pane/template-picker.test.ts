// @vitest-environment jsdom
// The Templates picker against the shipped taskpane.html: the rows the registry
// builds into it (a name, a line, an Insert button carrying the entry's id as
// its data-action), the search that filters them as you type, Enter inserting
// the first match, Escape clearing the query, and the empty state. The pane's
// own guard is replaced by a spy: what is proven here is which action a press
// asks for, not what the action does.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { TEMPLATE_GROUPS, TEMPLATES } from "../templates/index";
import { installTemplatePicker, renderTemplateList } from "./template-picker";

function paneMarkup(): string {
  return readFileSync(join(process.cwd(), "taskpane.html"), "utf8");
}

const search = (): HTMLInputElement =>
  document.getElementById("template-search") as HTMLInputElement;
const list = (): HTMLElement =>
  document.getElementById("template-list") as HTMLElement;
const emptyNote = (): HTMLElement =>
  document.getElementById("template-empty") as HTMLElement;

function type(query: string): void {
  search().value = query;
  search().dispatchEvent(new Event("input", { bubbles: true }));
}

function press(key: string): void {
  search().dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
}

function visibleActions(): string[] {
  return [
    ...list().querySelectorAll<HTMLButtonElement>(
      ".template-row:not([hidden]) button[data-action]",
    ),
  ].map((button) => button.dataset.action ?? "");
}

describe("the shipped Templates markup", () => {
  it("carries the search box, the list and the empty note, and no hand-written template button", () => {
    const doc = new DOMParser().parseFromString(paneMarkup(), "text/html");
    expect(doc.getElementById("template-search")).toBeInstanceOf(
      HTMLInputElement,
    );
    expect(doc.getElementById("template-list")).not.toBeNull();
    expect(doc.getElementById("template-empty")?.textContent?.trim()).toBe(
      "No template matches",
    );
    expect(
      doc.querySelectorAll('button[data-action^="template-"]'),
    ).toHaveLength(0);
  });

  it("keeps the section's heading id the help and the rig key on", () => {
    const doc = new DOMParser().parseFromString(paneMarkup(), "text/html");
    const section = doc.getElementById("templates-heading")?.closest("section");
    expect(section?.getAttribute("aria-labelledby")).toBe("templates-heading");
    expect(section?.querySelector("#template-search")).not.toBeNull();
  });
});

describe("renderTemplateList", () => {
  beforeEach(() => {
    document.body.innerHTML = paneMarkup();
  });

  it("builds one headed group per group that has a template, in the picker's order", () => {
    renderTemplateList(list(), TEMPLATES);
    const groups = [...list().querySelectorAll<HTMLElement>(".template-group")];
    const shown = TEMPLATE_GROUPS.filter((group) =>
      TEMPLATES.some((entry) => entry.group === group),
    );
    expect(groups.map((node) => node.dataset.group)).toEqual(shown);
    for (const group of groups) {
      const heading = group.querySelector("h3");
      expect(heading?.textContent).toBe(group.dataset.group);
      expect(group.getAttribute("aria-labelledby")).toBe(heading?.id);
    }
  });

  it("gives every entry one row with its name, its line and an Insert button", () => {
    renderTemplateList(list(), TEMPLATES);
    for (const entry of TEMPLATES) {
      const row = list().querySelector<HTMLElement>(
        `.template-row[data-template="${entry.id}"]`,
      );
      expect(row, entry.id).not.toBeNull();
      expect(row?.querySelector("strong")?.textContent).toBe(entry.name);
      expect(row?.textContent).toContain(entry.description);
      const button = row?.querySelector<HTMLButtonElement>("button");
      expect(button?.dataset.action).toBe(entry.id);
      expect(button?.type).toBe("button");
      expect(button?.textContent?.trim()).toBe("Insert");
      expect(button?.getAttribute("aria-label")).toBe(`Insert ${entry.name}`);
      // The tooltip keeps what the old hand-written buttons said: the block's
      // own long description and the size it will write.
      expect(button?.title).toContain(entry.block.description);
      expect(button?.title).toContain(
        `${String(entry.block.rows)} rows by ${String(entry.block.cols)} columns`,
      );
    }
  });

  it("puts each template under the group its entry names", () => {
    renderTemplateList(list(), TEMPLATES);
    for (const entry of TEMPLATES) {
      const row = list().querySelector(`[data-template="${entry.id}"]`);
      expect(row?.closest<HTMLElement>(".template-group")?.dataset.group).toBe(
        entry.group,
      );
    }
  });

  it("carries every data-action once, so search and sweep can tell the buttons apart", () => {
    renderTemplateList(list(), TEMPLATES);
    const actions = [
      ...list().querySelectorAll<HTMLElement>("button[data-action]"),
    ].map((button) => button.dataset.action);
    expect(actions).toHaveLength(TEMPLATES.length);
    expect(new Set(actions).size).toBe(TEMPLATES.length);
  });

  it("replaces what stood there when it renders again", () => {
    renderTemplateList(list(), TEMPLATES);
    renderTemplateList(list(), TEMPLATES);
    expect(list().querySelectorAll(".template-row")).toHaveLength(
      TEMPLATES.length,
    );
  });

  it("leaves out a group with nothing to show", () => {
    const valuation = TEMPLATES.filter((entry) => entry.group === "Valuation");
    renderTemplateList(list(), valuation);
    expect(
      [...list().querySelectorAll<HTMLElement>(".template-group")].map(
        (node) => node.dataset.group,
      ),
    ).toEqual(["Valuation"]);
  });

  it("builds into a document that is not the page's own", () => {
    const scratch = new DOMParser().parseFromString(
      "<div id='l'></div>",
      "text/html",
    );
    const target = scratch.getElementById("l") as HTMLElement;
    renderTemplateList(target, TEMPLATES);
    expect(target.querySelectorAll("button[data-action]")).toHaveLength(
      TEMPLATES.length,
    );
  });
});

describe("installTemplatePicker", () => {
  const run = vi.fn<(action: string) => void>();

  beforeEach(() => {
    run.mockReset();
    document.body.innerHTML = paneMarkup();
    installTemplatePicker(run);
  });

  it("lists every template at first, with the empty note hidden", () => {
    expect(visibleActions()).toEqual(TEMPLATES.map((entry) => entry.id));
    expect(emptyNote().hidden).toBe(true);
  });

  it("asks for an entry's action when its Insert button is pressed", () => {
    const dcf = list().querySelector<HTMLButtonElement>(
      'button[data-action="template-dcf"]',
    );
    dcf?.click();
    expect(run).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledWith("template-dcf");
  });

  it("asks for nothing when the press lands on a row but not on its button", () => {
    list()
      .querySelector<HTMLElement>(".template-row strong")
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(run).not.toHaveBeenCalled();
  });

  it("filters the rows as you type, by name, description or group", () => {
    type("bridge");
    expect(visibleActions()).toContain("template-ebitda-bridge");
    expect(visibleActions().length).toBeLessThan(TEMPLATES.length);
    type("waterfall");
    expect(visibleActions()).toContain("template-ebitda-bridge");
    type("lbo");
    expect(visibleActions()).toContain("template-debt-schedule");
    for (const action of visibleActions()) {
      expect(TEMPLATES.find((entry) => entry.id === action)?.group).toBe("LBO");
    }
  });

  it("hides a group whose rows are all filtered out", () => {
    type("lbo");
    const shown = [
      ...list().querySelectorAll<HTMLElement>(".template-group:not([hidden])"),
    ].map((node) => node.dataset.group);
    expect(shown).toEqual(["LBO"]);
  });

  it("says No template matches when nothing does, and drops the note again", () => {
    type("zzz");
    expect(visibleActions()).toEqual([]);
    expect(emptyNote().hidden).toBe(false);
    expect(emptyNote().textContent?.trim()).toBe("No template matches");
    type("");
    expect(emptyNote().hidden).toBe(true);
    expect(visibleActions()).toHaveLength(TEMPLATES.length);
  });

  it("inserts the first visible match when Enter is pressed in the search", () => {
    type("dcf");
    press("Enter");
    expect(run).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledWith("template-dcf");
  });

  it("takes the first of several matches, in the list's own order", () => {
    type("valuation");
    const first = visibleActions()[0];
    press("Enter");
    expect(run).toHaveBeenCalledWith(first);
    expect(first).toBe(
      TEMPLATES.find((entry) => entry.group === "Valuation")?.id,
    );
  });

  it("does nothing on Enter when no template matches", () => {
    type("zzz");
    press("Enter");
    expect(run).not.toHaveBeenCalled();
  });

  it("does nothing on Enter in an empty search, whitespace included", () => {
    press("Enter");
    expect(run).not.toHaveBeenCalled();
    type("   ");
    press("Enter");
    expect(run).not.toHaveBeenCalled();
    // A query that was typed and cleared again is empty too.
    type("dcf");
    type("");
    press("Enter");
    expect(run).not.toHaveBeenCalled();
  });

  it("marks what the query hides with data-filtered, and clears it with the query", () => {
    type("dcf");
    const hiddenRows = [
      ...list().querySelectorAll<HTMLElement>(".template-row"),
    ].filter((row) => row.hidden);
    expect(hiddenRows.length).toBeGreaterThan(0);
    expect(hiddenRows.every((row) => row.hasAttribute("data-filtered"))).toBe(
      true,
    );
    const shown = list().querySelector<HTMLElement>(
      '.template-row[data-template="template-dcf"]',
    );
    expect(shown?.hidden).toBe(false);
    expect(shown?.hasAttribute("data-filtered")).toBe(false);
    // A group with no row left carries the mark too; one with a row does not.
    const groups = [...list().querySelectorAll<HTMLElement>(".template-group")];
    for (const group of groups) {
      expect(group.hasAttribute("data-filtered"), group.dataset.group).toBe(
        group.hidden,
      );
    }
    expect(groups.some((group) => group.hidden)).toBe(true);

    type("");
    const marked = list().querySelectorAll("[data-filtered]");
    expect(marked.length).toBe(0);
  });

  it("does nothing on Enter while the pane is busy and its buttons are disabled", () => {
    for (const button of list().querySelectorAll("button")) {
      button.disabled = true;
    }
    type("dcf");
    press("Enter");
    expect(run).not.toHaveBeenCalled();
  });

  it("clears the query and brings every row back on Escape", () => {
    type("zzz");
    press("Escape");
    expect(search().value).toBe("");
    expect(visibleActions()).toHaveLength(TEMPLATES.length);
    expect(emptyNote().hidden).toBe(true);
  });

  it("leaves Escape alone when the box is already empty", () => {
    const event = new KeyboardEvent("keydown", {
      key: "Escape",
      bubbles: true,
      cancelable: true,
    });
    search().dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
  });

  it("reads a word at a time, so a longer query narrows the list further", () => {
    type("cash");
    const cash = visibleActions();
    type("cash payback");
    expect(visibleActions().length).toBeLessThanOrEqual(cash.length);
    expect(visibleActions()).toContain("template-npv-irr");
    expect(visibleActions()).not.toContain("template-sensitivity");
  });
});
