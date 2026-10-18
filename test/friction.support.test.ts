// @vitest-environment jsdom
// The friction action counter itself: what costs an action (a click, a key, a
// typed value, a paste) and what is free (a hover), what a finished flow
// records, that a flow over its target fails, and the scoreboard it prints.
import { beforeEach, describe, expect, it } from "vitest";
import { FRICTION_TARGETS } from "./friction-targets";
import { actions, begin, finish, scoreboard } from "./friction.support";

let field: HTMLInputElement;
let button: HTMLButtonElement;

beforeEach(() => {
  document.body.innerHTML =
    '<input id="field" /><button id="button">Go</button>';
  field = document.getElementById("field") as HTMLInputElement;
  button = document.getElementById("button") as HTMLButtonElement;
});

describe("the target table", () => {
  const rowsOf = (prefix: string): (string | number)[][] =>
    FRICTION_TARGETS.filter((row) => row.id.startsWith(prefix)).map(
      ({ id, flow, target }) => [id, flow, target],
    );

  it("holds the Excel rows of slices 1 and 2, then the style-cycle key, with their targets", () => {
    expect(rowsOf("excel-")).toEqual([
      ["excel-brand-fill", "Brand fill on cells", 1],
      ["excel-brand-font", "Brand font color on cells", 1],
      ["excel-brand-shade", "A shade of a brand color", 1],
      ["excel-palette-paste", "Palette from a brand sheet (8 colors)", 2],
      ["excel-style-save", "Save a named custom style", 3],
      ["excel-style-apply", "Apply a custom style", 1],
      ["excel-style-restyle", "Restyle a model after a palette change", 1],
      ["excel-style-cycle", "Step the selection through your styles", 1],
    ]);
  });

  it("holds the four PowerPoint brand-color rows of slice 3 with their targets", () => {
    const brand = rowsOf("ppt-").filter(
      ([id]) => !String(id).startsWith("ppt-style-"),
    );
    expect(brand).toEqual([
      ["ppt-brand-fill", "Brand color on the selected shapes", 1],
      ["ppt-theme-put", "Brand colors in PowerPoint's own picker", 1],
      [
        "ppt-palette-link-paste",
        "Palette into PowerPoint with a link paste",
        0,
      ],
      ["ppt-palette-paste", "Palette into PowerPoint without links", 3],
    ]);
  });

  it("holds the two shape-style rows of slice 4 with their targets", () => {
    expect(rowsOf("ppt-style-")).toEqual([
      ["ppt-style-save", "Save a shape style", 3],
      ["ppt-style-apply", "Apply a shape style to the selected shapes", 1],
    ]);
  });

  it("names every flow once", () => {
    const ids = FRICTION_TARGETS.map((row) => row.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe("what an action costs", () => {
  it("counts a click, a key, a typed value and a paste as one each", () => {
    begin("excel-brand-fill");
    let presses = 0;
    button.addEventListener("click", () => (presses += 1));
    actions.click(button);
    actions.key("Enter");
    actions.type(field, "a long value, typed in one go");
    actions.paste(field, "#B27E54");
    expect(() => finish("excel-brand-fill")).toThrow(/took 4 actions/);
    expect(presses).toBe(1);
  });

  it("counts a hover as nothing", () => {
    begin("excel-brand-fill");
    let seen = 0;
    button.addEventListener("mouseover", () => (seen += 1));
    actions.hover(button);
    actions.hover(button);
    actions.click(button);
    expect(finish("excel-brand-fill")).toBe(1);
    expect(seen).toBe(2);
  });

  it("counts an action made in another window as one, with the words it stands for", () => {
    begin("ppt-palette-paste");
    actions.elsewhere("Excel: Copy palette JSON");
    actions.click(button);
    actions.paste(field, "{}");
    expect(finish("ppt-palette-paste")).toBe(3);
  });

  it("names the other window's action when a flow runs over", () => {
    begin("ppt-brand-fill");
    actions.elsewhere("Excel: Copy palette JSON");
    actions.click(button);
    expect(() => finish("ppt-brand-fill")).toThrow(
      "took 2 actions (Excel: Copy palette JSON, in another window), the target is 1",
    );
  });

  it("gives a click the focus a pointer press gives", () => {
    begin("excel-brand-fill");
    actions.click(field);
    expect(document.activeElement).toBe(field);
    finish("excel-brand-fill");
  });

  it("types one value as input then change, and keeps it", () => {
    begin("excel-palette-paste");
    const seen: string[] = [];
    field.addEventListener("input", () => seen.push(`input:${field.value}`));
    field.addEventListener("change", () => seen.push(`change:${field.value}`));
    actions.type(field, "#112233");
    expect(seen).toEqual(["input:#112233", "change:#112233"]);
    finish("excel-palette-paste");
  });

  it("hands a paste's text to the field's clipboardData", () => {
    begin("excel-palette-paste");
    let text = "";
    field.addEventListener("paste", (event) => {
      text = (event as ClipboardEvent).clipboardData?.getData("text") ?? "";
    });
    actions.paste(field, "Gold #B27E54");
    expect(text).toBe("Gold #B27E54");
    finish("excel-palette-paste");
  });

  it("sends a chord with its modifiers to the focused element", () => {
    begin("excel-brand-fill");
    button.focus();
    const seen: string[] = [];
    button.addEventListener("keydown", (event) => {
      seen.push(
        `${event.ctrlKey ? "ctrl+" : ""}${event.shiftKey ? "shift+" : ""}${event.key}`,
      );
    });
    actions.key("Ctrl+Shift+F");
    expect(seen).toEqual(["ctrl+shift+F"]);
    finish("excel-brand-fill");
  });

  it("refuses a chord it cannot read", () => {
    begin("excel-brand-fill");
    expect(() => actions.key("Hyper+F")).toThrow(/bad chord/);
  });

  // Office hears an add-in shortcut, not the page: what the host would run on the key runs here.
  it("counts an Office shortcut as one action and runs what the host would run", () => {
    begin("excel-style-cycle");
    let ran = 0;
    actions.shortcut("Ctrl+Shift+Alt+1", () => (ran += 1));
    expect(ran).toBe(1);
    expect(finish("excel-style-cycle")).toBe(1);
  });

  it("refuses an Office shortcut chord it cannot read, before running anything", () => {
    begin("excel-style-cycle");
    let ran = 0;
    expect(() => actions.shortcut("Hyper+1", () => (ran += 1))).toThrow(
      /bad chord/,
    );
    expect(ran).toBe(0);
  });
});

describe("a flow", () => {
  it("passes at its target and records the count", () => {
    begin("excel-palette-paste");
    actions.click(field);
    actions.paste(field, "#B27E54");
    expect(finish("excel-palette-paste")).toBe(2);
    expect(scoreboard()).toMatch(
      /Palette from a brand sheet \(8 colors\)\s+2\s+2\s+ok/,
    );
  });

  it("fails above its target, naming the flow and both numbers", () => {
    begin("excel-brand-shade");
    actions.click(button);
    actions.click(button);
    expect(() => finish("excel-brand-shade")).toThrow(
      "A shade of a brand color took 2 actions, the target is 1",
    );
    expect(scoreboard()).toMatch(/A shade of a brand color\s+2\s+1\s+OVER/);
  });

  it("refuses an id with no target row, and an action with no flow begun", () => {
    expect(() => begin("not-a-flow")).toThrow("no target row");
    begin("excel-brand-fill");
    finish("excel-brand-fill");
    expect(() => actions.click(button)).toThrow("begin(id)");
    expect(() => finish("excel-brand-font")).toThrow("was not begun");
  });
});

describe("the scoreboard", () => {
  it("lists every target row, with flows not run marked so", () => {
    const table = scoreboard();
    for (const row of FRICTION_TARGETS) expect(table).toContain(row.flow);
    expect(table.split("\n")[0]).toMatch(/^Flow\s+Today\s+Target\s+Verdict$/);
    expect(table).toMatch(/Brand font color on cells\s+-\s+1\s+not run/);
  });
});
