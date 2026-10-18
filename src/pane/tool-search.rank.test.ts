// @vitest-environment jsdom
// Find a tool over the SHIPPED Tools, Workbook, Links and Brand tabs: what the
// four queries the usability pass named now find, the badge a row carries, a
// long sentence clipped at a word, and the one line a query with no hit gets.
// Loaded like tool-search.test.ts: the markup first, then the module, fresh per
// test (./shared builds the tabs and the toast off whatever DOM is live).
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { TEMPLATES } from "../templates/index";
import { renderTemplateList } from "./template-picker";

async function load() {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  renderTemplateList(
    document.getElementById("template-list") as HTMLElement,
    TEMPLATES,
  );
  const module = await import("./tool-search");
  module.installToolSearch(document);
  return module;
}

const input = (): HTMLInputElement =>
  document.getElementById("tool-search") as HTMLInputElement;
const results = (): HTMLElement =>
  document.getElementById("tool-search-results") as HTMLElement;
const emptyLine = (): HTMLElement =>
  document.getElementById("tool-search-empty") as HTMLElement;

function type(value: string): void {
  input().value = value;
  input().dispatchEvent(new Event("input"));
}

function found(query: string): (string | undefined)[] {
  type(query);
  return Array.from(results().querySelectorAll("li")).map(
    (li) => li.dataset.toolAction,
  );
}

describe("the catalogue carries each row's section", () => {
  it("reads the h2 and the kicker the button sits under", async () => {
    const { toolEntries } = await load();
    const entries = toolEntries(document);
    const of = (action: string) => entries.find((e) => e.action === action);

    expect(of("export-selection")).toMatchObject({
      heading: "Export",
      kicker: "Send to PowerPoint",
    });
    expect(of("style-save")).toMatchObject({
      heading: "Custom styles",
      kicker: "Your styles",
    });
    expect(of("undo")).toMatchObject({
      heading: "Model tools",
      kicker: "Speed",
    });
  });

  it("keeps the help card's own words out of it: the ? toggle is not part of the heading", async () => {
    const { toolEntries } = await load();
    const { installHelp } = await import("../ui/help");
    installHelp(document);
    const entry = toolEntries(document).find((e) => e.action === "style-save");
    expect(entry?.heading).toBe("Custom styles");
  });
});

describe("the four queries the usability pass named, on the shipped pane", () => {
  it('"multiple" lists the 0.0x cycle first, not a template that writes a block', async () => {
    await load();
    const actions = found("multiple");
    expect(actions[0]).toBe("cycle-number-multiple");
    expect(actions.indexOf("template-trading-comps")).toBeGreaterThan(0);
  });

  it('"send to powerpoint" lists the four Links export buttons first', async () => {
    await load();
    expect(found("send to powerpoint").slice(0, 4)).toEqual([
      "export-selection",
      "export-chart",
      "export-text",
      "export-picture",
    ]);
  });

  it('"trace precedents" lists Precedents first', async () => {
    await load();
    const actions = found("trace precedents");
    expect(actions[0]).toBe("trace-precedents");
    expect(actions).toContain("trace-precedents-all");
    expect(actions).not.toContain("trace-dependents");
  });

  it('"custom styles" lists the Custom styles section\'s Save style before Scan styles', async () => {
    await load();
    const actions = found("custom styles");
    expect(actions).toContain("style-save");
    expect(actions).toContain("styles-scan");
    expect(actions.indexOf("style-save")).toBeLessThan(
      actions.indexOf("styles-scan"),
    );
  });

  it("puts every template Insert row after the other rows it ties with", async () => {
    await load();
    const actions = found("insert");
    const first = actions.findIndex((a) => a?.startsWith("template-"));
    expect(first).toBeGreaterThan(0);
    expect(
      actions.slice(0, first).some((a) => a?.startsWith("template-")),
    ).toBe(false);
    expect(actions.slice(0, 2)).toEqual(["insert-color-key", "insert-toc"]);
  });

  it("still finds undo first for undo", async () => {
    await load();
    expect(found("undo")[0]).toBe("undo");
  });
});

describe("the tab badge on a row", () => {
  it("names the tab the control sits on, Edit colors included, because Enter takes you to the control", async () => {
    await load();
    type("edit colors");
    const badge = results().querySelector(".tool-search-tab");
    expect(badge?.textContent).toBe("Tools");
  });
});

describe("a long sentence in a row", () => {
  it("is cut at a word boundary, never in the middle of one", async () => {
    await load();
    type("Insert Trading comps");
    const small = results().querySelector("li small");
    expect(small?.textContent).toBe(
      "Writes a trading comps table: six peers in, three multiples, six statistics and an…",
    );
  });

  it("is cut at a word in every row of a broad query", async () => {
    const { toolEntries } = await load();
    const sentences = new Map(
      toolEntries(document).map((e) => [e.action, e.sentence]),
    );
    type("selection");
    const rows = Array.from(results().querySelectorAll("li"));
    expect(rows.length).toBeGreaterThan(3);
    for (const row of rows) {
      const shown = row.querySelector("small")?.textContent ?? "";
      const sentence = sentences.get(row.dataset.toolAction ?? "") ?? "";
      if (shown === sentence) continue;
      const kept = shown.slice(0, -1);
      expect(shown.endsWith("…")).toBe(true);
      expect(sentence.startsWith(kept)).toBe(true);
      expect(sentence[kept.length]).toMatch(/[\s,.;:]/);
    }
  });
});

describe("a query with no hit", () => {
  it("says so in one line that names the query and offers words that work", async () => {
    await load();
    type("frobnicate");

    expect(results().hidden).toBe(true);
    expect(emptyLine().hidden).toBe(false);
    expect(emptyLine().textContent?.trim()).toBe(
      'No tool matches "frobnicate". Try export, undo, color or dcf.',
    );
    expect(emptyLine().getAttribute("role")).toBe("status");
  });

  it("shows the query trimmed, and nothing for a blank one", async () => {
    await load();
    type("  frobnicate ");
    expect(emptyLine().textContent?.trim()).toBe(
      'No tool matches "frobnicate". Try export, undo, color or dcf.',
    );

    type("   ");
    expect(emptyLine().hidden).toBe(true);
  });

  it("goes the moment something matches, on clearing and on Escape", async () => {
    await load();
    type("frobnicate");
    expect(emptyLine().hidden).toBe(false);

    type("undo");
    expect(emptyLine().hidden).toBe(true);
    expect(results().hidden).toBe(false);

    type("frobnicate");
    type("");
    expect(emptyLine().hidden).toBe(true);

    type("frobnicate");
    input().dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
    expect(emptyLine().hidden).toBe(true);
    expect(input().value).toBe("");
  });

  it("is not shown over a list that has rows", async () => {
    await load();
    type("export");
    expect(results().hidden).toBe(false);
    expect(emptyLine().hidden).toBe(true);
  });

  it("speaks of a button the pane would not offer as no match, like any other miss", async () => {
    await load();
    // Copy report is disabled until a model check has run: not offered.
    type("copy report");
    expect(results().hidden).toBe(true);
    expect(emptyLine().textContent).toContain('No tool matches "copy report"');
  });
});
