// @vitest-environment jsdom
// installToolSearch pulls in ./shared, whose top-level installTabs/createToast
// calls need #tab-bar and #toast to already exist - and by the time any
// static top-level import runs, a beforeEach has not fired yet. So every
// test loads the real markup first and only then dynamically imports the
// module under test, with the module registry reset in between so each test
// sees a fresh catalogue built off its own fresh DOM.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { TEMPLATES } from "../templates/index";
import { installTemplatePicker, renderTemplateList } from "./template-picker";

// The shipped markup plus what the pane builds at boot before the catalogue is
// read: the Templates picker's Insert buttons (src/main.ts), which the tool
// search finds like any other button.
function paneRoot(): Document {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  renderTemplateList(
    document.getElementById("template-list") as HTMLElement,
    TEMPLATES,
  );
  return document;
}

// No explicit return type: annotating it would need an `import()` type,
// which the lint config forbids in favour of a top-level type import - and
// none of this module's exported types are needed here.
async function load() {
  vi.resetModules();
  paneRoot();
  return import("./tool-search");
}

function input(): HTMLInputElement {
  return document.getElementById("tool-search") as HTMLInputElement;
}

function results(): HTMLElement {
  return document.getElementById("tool-search-results") as HTMLElement;
}

function type(value: string): void {
  const el = input();
  el.value = value;
  el.dispatchEvent(new Event("input"));
}

function press(key: string): void {
  input().dispatchEvent(new KeyboardEvent("keydown", { key }));
}

// What the Templates picker's own box does when a modeller types in it.
function filterTemplatesBy(query: string): void {
  const box = document.getElementById("template-search") as HTMLInputElement;
  box.value = query;
  box.dispatchEvent(new Event("input"));
}

function resultActions(): (string | undefined)[] {
  return Array.from(results().querySelectorAll("li")).map(
    (li) => li.dataset.toolAction,
  );
}

describe("toolEntries", () => {
  it("finds a known button on its own tab", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const entry = entries.find((e) => e.action === "export-selection");
    expect(entry).toBeDefined();
    expect(entry?.label).toBe("Export selection");
    expect(entry?.tab).toBe("Links");
    expect(entry?.sentence.length).toBeGreaterThan(0);
  });

  it("leaves out a button that is hidden inside its own panel", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    expect(entries.some((e) => e.action === "styles-delete")).toBe(false);
  });

  it("catalogues the three narrow pastes under their own labels", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const labels = new Map(entries.map((e) => [e.action, e.label]));
    expect(labels.get("paste-duplicate")).toBe("Paste: duplicate formulas");
    expect(labels.get("paste-number-formats")).toBe("Paste: number formats");
    expect(labels.get("paste-row-heights")).toBe("Paste: row heights");
    for (const action of [
      "paste-duplicate",
      "paste-number-formats",
      "paste-row-heights",
    ]) {
      const entry = entries.find((e) => e.action === action);
      expect(entry?.tab).toBe("Tools");
      expect(entry?.sentence.length).toBeGreaterThan(0);
    }
  });

  it("still catalogues a button on a tab that is not the active one", async () => {
    const { toolEntries } = await load();
    // Workbook, Links and Brand all start hidden; only Tools is active.
    document.getElementById("view-links")!.hidden = true;
    const entries = toolEntries(document);
    expect(entries.some((e) => e.action === "export-selection")).toBe(true);
  });

  it("catalogues the hygiene cycles and the sheet tools on their own tabs", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const found = (action: string) =>
      entries.find((entry) => entry.action === action);

    expect(found("cycle-indent")).toMatchObject({
      label: "Indent",
      tab: "Tools",
    });
    expect(found("cycle-align")?.tab).toBe("Tools");
    expect(found("cycle-underline")?.tab).toBe("Tools");
    expect(found("sheets-unhide-all")).toMatchObject({
      label: "Unhide all",
      tab: "Workbook",
    });
    expect(found("sheets-move-end")?.tab).toBe("Workbook");
    expect(found("clean-past-data")).toMatchObject({
      label: "Clean past the data",
      tab: "Workbook",
    });
    // Every one of them carries its help sentence into the search row.
    for (const action of [
      "cycle-indent",
      "cycle-align",
      "cycle-underline",
      "sheets-unhide-all",
      "sheets-show-only",
      "sheets-bury",
      "sheets-move-up",
      "sheets-move-down",
      "sheets-move-end",
      "clean-past-data",
    ]) {
      expect([action, found(action)?.sentence.length ?? 0]).not.toEqual([
        action,
        0,
      ]);
    }
  });

  // Enter takes you to the control, so the badge names the tab the control sits
  // on, even for a button whose press moves the pane on (Edit colors opens Brand).
  it("lists a button that moves the pane on under the tab it sits on", async () => {
    const { toolEntries } = await load();
    const entry = toolEntries(document).find(
      (e) => e.action === "edit-brand-colors",
    );
    expect(entry?.label).toBe("Edit colors");
    // The button sits in the Tools tab's Brand colors section.
    expect(
      document.getElementById("edit-brand-colors")?.closest("#view-tools"),
    ).not.toBeNull();
    expect(entry?.tab).toBe("Tools");
  });

  it("labels a glyph-only button by its aria-label, not the glyph on its face", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const entry = entries.find((e) => e.action === "refresh-selection");
    expect(entry?.label).toBe("Refresh selection");
  });
});

// What Enter and a click on a result do (go to the control, never press it) is
// in tool-search.go.test.ts; the ranking and the words are in
// tool-search.rank.test.ts.
describe("installToolSearch", () => {
  it("lists a button when its name is typed", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");

    expect(results().hidden).toBe(false);
    expect(results().textContent).toContain("Export selection");
  });

  it("Escape clears the box and hides the list", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("export");
    expect(results().hidden).toBe(false);

    press("Escape");

    expect(input().value).toBe("");
    expect(results().hidden).toBe(true);
  });

  it("shows the hint only while empty and focused", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    const hint = document.getElementById("tool-search-hint")!;

    input().focus();
    expect(hint.hidden).toBe(false);

    type("export");
    expect(hint.hidden).toBe(true);

    type("");
    expect(hint.hidden).toBe(false);

    input().blur();
    expect(hint.hidden).toBe(true);
  });

  it("ArrowDown/ArrowUp move the active row, clamped at both ends", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    // "waterfall" matches the Waterfall button by label and the EBITDA
    // bridge template by its help sentence, giving more than one row.
    type("waterfall");
    const rows = () => Array.from(results().querySelectorAll("li"));
    expect(rows().length).toBeGreaterThan(1);
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    press("ArrowDown");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("false");
    expect(rows()[1]!.getAttribute("aria-selected")).toBe("true");
    expect(input().getAttribute("aria-activedescendant")).toBe(rows()[1]!.id);

    press("ArrowUp");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    // Clamped, not wrapped: one more ArrowUp at the top stays put.
    press("ArrowUp");
    expect(rows()[0]!.getAttribute("aria-selected")).toBe("true");

    for (let i = 0; i < rows().length + 2; i += 1) press("ArrowDown");
    const last = rows().length - 1;
    expect(rows()[last]!.getAttribute("aria-selected")).toBe("true");
  });

  it('"/" focuses the box from elsewhere in the pane, but not from a field', async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    const other = document.getElementById(
      "refresh-selection",
    ) as HTMLButtonElement;
    other.focus();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "/" }));
    expect(document.activeElement).toBe(input());

    // Typing "/" while already inside a text field must not be hijacked.
    const findQuery = document.getElementById("find-query") as HTMLInputElement;
    findQuery.focus();
    document.dispatchEvent(new KeyboardEvent("keydown", { key: "/" }));
    expect(document.activeElement).toBe(findQuery);
  });

  it("does not offer a disabled button as a result", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    // copy-model-check ("Copy report") starts disabled until a model check
    // has run, and is the only catalogued button naming both words.
    type("copy report");

    expect(results().hidden).toBe(true);
  });

  it("finds a glyph-only button by its aria-label", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    type("refresh selection");

    expect(results().hidden).toBe(false);
    expect(results().textContent).toContain("Refresh selection");
  });

  it("does not offer the New project prompt's Create/Cancel before New project opens it", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);

    // #project-prompt (holding project-ok "Create" and project-cancel
    // "Cancel") starts hidden; New project itself does not and stays a
    // valid result for the same query.
    type("create");
    const actions = Array.from(results().querySelectorAll("li")).map(
      (li) => li.dataset.toolAction,
    );

    expect(actions).toContain("new-project");
    expect(actions).not.toContain("project-ok");
  });

  // The Templates search hides rows for its own query; that is not a reason
  // for Find a tool to refuse them, unlike the prompt above, which is closed.
  it("still offers a template the Templates search is hiding", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    installTemplatePicker(vi.fn());
    filterTemplatesBy("comps");
    const row = document.querySelector<HTMLElement>(
      '[data-template="template-wacc"]',
    );
    expect(row?.hidden).toBe(true);
    expect(row?.closest<HTMLElement>(".template-group")?.hidden).toBe(false);

    type("Insert WACC");
    expect(resultActions()[0]).toBe("template-wacc");
  });

  it("still offers a template whose whole group the Templates search is hiding", async () => {
    const { installToolSearch } = await load();
    installToolSearch(document);
    installTemplatePicker(vi.fn());
    filterTemplatesBy("dcf");
    const group = document.querySelector<HTMLElement>(
      '.template-group[data-group="LBO"]',
    );
    expect(group?.hidden).toBe(true);

    type("Insert Debt schedule");
    expect(resultActions()[0]).toBe("template-debt-schedule");
  });
});
