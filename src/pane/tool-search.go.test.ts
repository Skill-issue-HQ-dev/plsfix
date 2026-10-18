// @vitest-environment jsdom
// Enter, or a click, on a Find a tool result takes you to the tool and stops
// there: its tab, the control in view and in focus, a short mark on it. It never
// presses the control: a second Enter, now on the focused control, does that the
// normal way. Over the shipped markup; the booted pane is in
// test/tool-search.enter.integration.test.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { pressEnterOn } from "../../test/native-enter";
import { TEMPLATES } from "../templates/index";
import { installTemplatePicker, renderTemplateList } from "./template-picker";

const scroll = vi.fn();

afterEach(() => {
  vi.useRealTimers();
  Reflect.deleteProperty(HTMLElement.prototype, "scrollIntoView");
  scroll.mockClear();
});

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
  // jsdom has no layout, so no scrollIntoView: a stub says it was asked.
  Object.defineProperty(HTMLElement.prototype, "scrollIntoView", {
    value: scroll,
    configurable: true,
  });
  const module = await import("./tool-search");
  module.installToolSearch(document);
  return module;
}

const input = (): HTMLInputElement =>
  document.getElementById("tool-search") as HTMLInputElement;
const results = (): HTMLElement =>
  document.getElementById("tool-search-results") as HTMLElement;
const control = (action: string): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>(
    `button[data-action="${action}"]`,
  ) ?? (document.getElementById(action) as HTMLButtonElement);

function type(value: string): void {
  input().value = value;
  input().dispatchEvent(new Event("input"));
}

function press(key: string): void {
  input().dispatchEvent(
    new KeyboardEvent("keydown", { key, cancelable: true }),
  );
}

describe("Enter on a result", () => {
  it("shows the control's tab, puts focus on the control and does not press it", async () => {
    await load();
    const button = control("export-selection");
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    type("export");
    press("Enter");

    expect(document.getElementById("view-links")!.hidden).toBe(false);
    expect(document.getElementById("view-tools")!.hidden).toBe(true);
    expect(
      document.getElementById("tab-links")!.getAttribute("aria-selected"),
    ).toBe("true");
    expect(document.activeElement).toBe(button);
    expect(onClick).not.toHaveBeenCalled();
  });

  it("clears the box and closes the list, so the pane is not left pushed down", async () => {
    await load();
    type("export");
    expect(results().hidden).toBe(false);
    press("Enter");
    expect(input().value).toBe("");
    expect(results().hidden).toBe(true);
    expect(document.getElementById("tool-search-hint")!.hidden).toBe(true);
  });

  it("keeps the browser's own Enter away from the box", async () => {
    await load();
    type("export");
    const keydown = new KeyboardEvent("keydown", {
      key: "Enter",
      cancelable: true,
    });
    input().dispatchEvent(keydown);
    // Focus moved to the control inside this keydown: left alone, the press
    // would follow it and the control would be pressed by the first Enter.
    expect(keydown.defaultPrevented).toBe(true);
  });

  it("scrolls the control into view", async () => {
    await load();
    type("export");
    press("Enter");
    expect(scroll).toHaveBeenCalledTimes(1);
    expect(scroll).toHaveBeenCalledWith({ block: "center" });
    expect(scroll.mock.contexts[0]).toBe(control("export-selection"));
  });

  it("works where the page has no scrollIntoView at all", async () => {
    await load();
    Reflect.deleteProperty(HTMLElement.prototype, "scrollIntoView");
    type("export");
    expect(() => press("Enter")).not.toThrow();
    expect(document.activeElement).toBe(control("export-selection"));
  });

  it("marks the control with the focus style for two seconds, then lets it go", async () => {
    vi.useFakeTimers();
    await load();
    type("export");
    press("Enter");
    const button = control("export-selection");
    expect(button.classList.contains("tool-found")).toBe(true);

    vi.advanceTimersByTime(1_999);
    expect(button.classList.contains("tool-found")).toBe(true);
    vi.advanceTimersByTime(2);
    expect(button.classList.contains("tool-found")).toBe(false);
  });

  it("lets the mark go the moment focus moves on", async () => {
    await load();
    type("export");
    press("Enter");
    const button = control("export-selection");
    expect(button.classList.contains("tool-found")).toBe(true);
    control("export-chart").focus();
    expect(button.classList.contains("tool-found")).toBe(false);
  });

  it("marks one control at a time", async () => {
    await load();
    type("export selection");
    press("Enter");
    type("push all");
    press("Enter");
    expect(control("export-selection").classList.contains("tool-found")).toBe(
      false,
    );
    expect(control("push-all").classList.contains("tool-found")).toBe(true);
    expect(document.querySelectorAll(".tool-found")).toHaveLength(1);
  });

  it("goes to the arrow-selected row, not the first", async () => {
    await load();
    type("waterfall");
    const second = results().querySelectorAll("li")[1]!.dataset.toolAction!;
    expect(second).not.toBe(
      results().querySelectorAll("li")[0]!.dataset.toolAction,
    );
    press("ArrowDown");
    press("Enter");
    expect(document.activeElement).toBe(control(second));
  });

  it("goes to Edit colors where it sits, on Tools, and does not open the Brand tab", async () => {
    await load();
    type("edit colors");
    expect(results().querySelector(".tool-search-tab")?.textContent).toBe(
      "Tools",
    );
    const opened = vi.fn();
    control("edit-brand-colors").addEventListener("click", opened);
    press("Enter");

    expect(
      document.getElementById("tab-tools")!.getAttribute("aria-selected"),
    ).toBe("true");
    expect(document.activeElement).toBe(control("edit-brand-colors"));
    expect(opened).not.toHaveBeenCalled();
  });

  it("does nothing for a control that went disabled after the list drew", async () => {
    await load();
    type("export");
    const button = control("export-selection");
    button.disabled = true;
    press("Enter");
    expect(document.activeElement).not.toBe(button);
    expect(input().value).toBe("export");
  });

  it("does nothing at all when nothing matched", async () => {
    await load();
    type("frobnicate");
    const keydown = new KeyboardEvent("keydown", {
      key: "Enter",
      cancelable: true,
    });
    input().dispatchEvent(keydown);
    expect(keydown.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(document.body);
  });
});

describe("a second Enter, on the control the first one went to", () => {
  it("presses it the normal way, once", async () => {
    await load();
    const button = control("export-selection");
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    type("export");
    press("Enter");
    expect(onClick).not.toHaveBeenCalled();

    const target = document.activeElement as HTMLElement;
    const keydown = pressEnterOn(target);
    expect(keydown.defaultPrevented).toBe(false);
    expect(onClick).toHaveBeenCalledTimes(1);
  });
});

describe("a click on a result row", () => {
  it("goes to the control the way Enter does, without pressing it", async () => {
    await load();
    const button = control("undo");
    const onClick = vi.fn();
    button.addEventListener("click", onClick);

    type("undo");
    results()
      .querySelector("li")!
      .dispatchEvent(new MouseEvent("click", { bubbles: true }));

    expect(document.activeElement).toBe(button);
    expect(button.classList.contains("tool-found")).toBe(true);
    expect(onClick).not.toHaveBeenCalled();
    expect(input().value).toBe("");
    expect(results().hidden).toBe(true);
  });
});

describe("a template the Templates search is hiding", () => {
  function filterTemplatesBy(query: string): void {
    const box = document.getElementById("template-search") as HTMLInputElement;
    box.value = query;
    box.dispatchEvent(new Event("input"));
  }

  it("is brought back by taking that search's words away, and gone to, not inserted", async () => {
    await load();
    const run = vi.fn();
    installTemplatePicker(run);
    filterTemplatesBy("comps");
    const row = document.querySelector<HTMLElement>(
      '[data-template="template-wacc"]',
    )!;
    expect(row.hidden).toBe(true);

    type("Insert WACC");
    press("Enter");

    expect(
      (document.getElementById("template-search") as HTMLInputElement).value,
    ).toBe("");
    expect(row.hidden).toBe(false);
    expect(row.closest<HTMLElement>(".template-group")!.hidden).toBe(false);
    expect(document.activeElement).toBe(control("template-wacc"));
    expect(run).not.toHaveBeenCalled();
  });

  it("leaves the Templates search alone for a template it is showing", async () => {
    await load();
    installTemplatePicker(vi.fn());
    filterTemplatesBy("wacc");
    type("Insert WACC");
    press("Enter");
    expect(
      (document.getElementById("template-search") as HTMLInputElement).value,
    ).toBe("wacc");
    expect(document.activeElement).toBe(control("template-wacc"));
  });

  it("then inserts it on a second Enter, through the picker's own press", async () => {
    await load();
    const run = vi.fn();
    installTemplatePicker(run);
    filterTemplatesBy("comps");

    type("Insert WACC");
    press("Enter");
    pressEnterOn(document.activeElement as HTMLElement);

    expect(run).toHaveBeenCalledTimes(1);
    expect(run).toHaveBeenCalledWith("template-wacc");
  });
});

describe("the mark", () => {
  it("is the focus ring every control already has: 2 px, the link token, no color of its own", () => {
    const css = readFileSync(
      join(process.cwd(), "src/styles/search.css"),
      "utf8",
    );
    const rule = /\.tool-found\s*\{([^}]*)\}/.exec(css)?.[1] ?? "";
    expect(rule).toContain("outline: 2px solid var(--link)");
    expect(rule).not.toMatch(/#[0-9a-f]{3,8}\b/i);
  });
});
