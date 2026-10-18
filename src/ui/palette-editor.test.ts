// @vitest-environment jsdom
// The palette editor both panes share: the brand rows past Primary and Accent,
// an Add color button and a Paste colors field with its status line. Driven
// through events; the host is two spies (onPreview, onCommit), never storage.
import { beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import { brandColors, MAX_BRAND_COLORS } from "../palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "../settings";
import { applyBrandColors, renderPaletteEditor } from "./palette-editor";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const five = palette(GOLD, MARBLE, GREEN, SLATE, NEON);
const ten = palette(
  ...Array.from(
    { length: 10 },
    (_, at) => `#${String(at + 1).padStart(2, "0")}A0C0`,
  ),
);

type Spy = Mock<(settings: BrandSettings) => void>;
let container: HTMLElement;
let onPreview: Spy;
let onCommit: Spy;

function render(settings: BrandSettings = five): void {
  renderPaletteEditor(container, {
    settings,
    onPreview: (next) => onPreview(next),
    onCommit: (next) => onCommit(next),
  });
}

beforeEach(() => {
  document.body.innerHTML = '<div id="editor"></div>';
  container = document.getElementById("editor")!;
  onPreview = vi.fn<(settings: BrandSettings) => void>();
  onCommit = vi.fn<(settings: BrandSettings) => void>();
});

const rows = (): HTMLElement[] => [
  ...container.querySelectorAll<HTMLElement>("[data-brand-index]"),
];
const picker = (at: number): HTMLInputElement =>
  rows()[at]!.querySelector<HTMLInputElement>('input[type="color"]')!;
const hexField = (at: number): HTMLInputElement =>
  rows()[at]!.querySelector<HTMLInputElement>('input[type="text"]')!;
const removeButton = (at: number): HTMLButtonElement =>
  rows()[at]!.querySelector<HTMLButtonElement>("button")!;
const addButton = (): HTMLButtonElement =>
  container.querySelector<HTMLButtonElement>("#add-brand-color")!;
const pasteField = (): HTMLInputElement =>
  container.querySelector<HTMLInputElement>("#paste-brand-colors")!;
const status = (): HTMLElement =>
  container.querySelector<HTMLElement>("[data-palette-status]")!;

function fire(target: HTMLElement, type: string): void {
  target.dispatchEvent(new Event(type, { bubbles: true }));
}

function setValue(field: HTMLInputElement, value: string): void {
  field.value = value;
}

function paste(text: string): Event {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.assign(event, { clipboardData: { getData: () => text } });
  pasteField().dispatchEvent(event);
  return event;
}

function lastCall(spy: Spy): BrandSettings {
  return spy.mock.calls.at(-1)![0];
}

describe("what the editor draws", () => {
  it("shows one row per brand color past primary and accent", () => {
    render();
    expect(rows().map((row) => row.dataset.brandIndex)).toEqual([
      "2",
      "3",
      "4",
    ]);
    expect(
      rows().map((row) => row.querySelector("strong")?.textContent),
    ).toEqual(["Brand 3", "Brand 4", "Brand 5"]);
  });

  it("shows no row for a palette of two", () => {
    render(DEFAULT_SETTINGS);
    expect(rows()).toHaveLength(0);
    expect(addButton()).not.toBeNull();
  });

  it("gives each row a color picker, a hex field and a remove button", () => {
    render();
    expect(picker(0).value).toBe(GREEN.toLowerCase());
    expect(hexField(0).value).toBe(GREEN);
    expect(hexField(0).classList.contains("hex-input")).toBe(true);
    expect(hexField(0).maxLength).toBe(7);
    expect(picker(0).getAttribute("aria-label")).toBe("Brand 3 color picker");
    expect(hexField(0).getAttribute("aria-label")).toBe("Brand 3 hex value");
    expect(removeButton(0).getAttribute("aria-label")).toBe("Remove Brand 3");
    expect(removeButton(0).title).toBe("Remove Brand 3");
    expect(removeButton(0).type).toBe("button");
  });

  it("draws Add color, a Paste colors field and a polite status line", () => {
    render();
    expect(addButton().textContent).toBe("Add color");
    expect(addButton().type).toBe("button");
    expect(pasteField().getAttribute("aria-label")).toBe("Paste colors");
    expect(pasteField().placeholder).not.toBe("");
    expect(status().getAttribute("aria-live")).toBe("polite");
    expect(status().textContent).toBe("");
  });

  it("holds ten colors as eight rows", () => {
    render(ten);
    expect(rows()).toHaveLength(8);
  });
});

describe("a color picker", () => {
  it("previews on input and commits on change", () => {
    render();
    setValue(picker(1), "#abcdef");
    fire(picker(1), "input");
    expect(onCommit).not.toHaveBeenCalled();
    expect(lastCall(onPreview).brand).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      "#ABCDEF",
      NEON,
    ]);
    expect(hexField(1).value).toBe("#ABCDEF");

    fire(picker(1), "change");
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(lastCall(onCommit).brand).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      "#ABCDEF",
      NEON,
    ]);
  });

  it("leaves primary and accent where they were", () => {
    render();
    setValue(picker(0), "#abcdef");
    fire(picker(0), "change");
    expect(lastCall(onCommit).primary).toBe(GOLD);
    expect(lastCall(onCommit).accent).toBe(MARBLE);
  });
});

describe("a hex field", () => {
  it("previews a typed color that is complete and says nothing for half of one", () => {
    render();
    setValue(hexField(0), "#B27E5");
    fire(hexField(0), "input");
    expect(onPreview).not.toHaveBeenCalled();
    setValue(hexField(0), "#112233");
    fire(hexField(0), "input");
    expect(lastCall(onPreview).brand[2]).toBe("#112233");
    expect(picker(0).value).toBe("#112233");
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("commits on change, and once for Enter followed by change", () => {
    render();
    setValue(hexField(0), "112233");
    hexField(0).dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    fire(hexField(0), "change");
    expect(onCommit).toHaveBeenCalledTimes(1);
    expect(lastCall(onCommit).brand[2]).toBe("#112233");
  });

  it("says so and puts the old color back when the text is not a color", () => {
    render();
    setValue(hexField(0), "not a color");
    fire(hexField(0), "change");
    expect(onCommit).not.toHaveBeenCalled();
    expect(status().textContent).toBe(
      '"not a color" is not a hex color like #6D7E12.',
    );
    expect(hexField(0).value).toBe(GREEN);
  });
});

describe("remove", () => {
  it("takes that color out of the list and commits", () => {
    render();
    removeButton(1).click();
    expect(lastCall(onCommit).brand).toEqual([GOLD, MARBLE, GREEN, NEON]);
    expect(status().textContent).toBe(`Removed ${SLATE}.`);
  });

  it("leaves primary and accent in every case", () => {
    render();
    removeButton(0).click();
    removeButton(1).click();
    removeButton(2).click();
    for (const [settings] of onCommit.mock.calls) {
      expect(brandColors(settings).slice(0, 2)).toEqual([GOLD, MARBLE]);
    }
  });
});

describe("Add color", () => {
  it("appends one color the list does not hold yet, and commits", () => {
    render();
    addButton().click();
    const next = lastCall(onCommit);
    expect(next.brand).toHaveLength(6);
    expect(next.brand.slice(0, 5)).toEqual(five.brand);
    expect(next.brand[5]).toMatch(/^#[0-9A-F]{6}$/);
    expect(five.brand).not.toContain(next.brand[5]);
  });

  it("works on a palette of two", () => {
    render(DEFAULT_SETTINGS);
    addButton().click();
    expect(lastCall(onCommit).brand).toHaveLength(3);
  });

  it("says ten is the most, and adds nothing, at ten colors", () => {
    render(ten);
    addButton().click();
    expect(onCommit).not.toHaveBeenCalled();
    expect(status().textContent).toBe(
      `Ten colors is the most, so none was added. Remove one first.`,
    );
  });
});

describe("Paste colors", () => {
  const sheet =
    "Brand Gold #B27E54 Pantone 729 / Marble #282623 / Dark Green 273631 / rgb(65, 86, 79) / Neon #DEF970";

  it("replaces the whole list, primary and accent included, for two or more colors", () => {
    render(DEFAULT_SETTINGS);
    const event = paste(sheet);
    expect(event.defaultPrevented).toBe(true);
    const next = lastCall(onCommit);
    expect(next.brand).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
    expect(next.primary).toBe(GOLD);
    expect(next.accent).toBe(MARBLE);
    expect(status().textContent).toBe("5 colors read.");
    expect(pasteField().value).toBe("");
  });

  it("says how many more did not fit past ten", () => {
    render();
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    paste(twelve.join(" "));
    expect(lastCall(onCommit).brand).toHaveLength(MAX_BRAND_COLORS);
    expect(status().textContent).toBe("10 colors read, 2 more did not fit.");
  });

  it("appends a single color to the list", () => {
    render();
    paste("our accent is #123456");
    expect(lastCall(onCommit).brand).toEqual([...five.brand, "#123456"]);
    expect(status().textContent).toBe("Added #123456.");
  });

  it("refuses a single color politely at ten", () => {
    render(ten);
    paste("#123456");
    expect(onCommit).not.toHaveBeenCalled();
    expect(status().textContent).toBe(
      "Ten colors is the most, so #123456 was not added. Remove one first.",
    );
  });

  it("says when the text holds no color, and changes nothing", () => {
    render();
    paste("our brand is warm");
    expect(onCommit).not.toHaveBeenCalled();
    expect(status().textContent).toBe("No colors found in that text.");
  });

  it("reads palette JSON the way Copy palette JSON writes it", () => {
    render(DEFAULT_SETTINGS);
    paste(JSON.stringify({ brand: [GOLD, MARBLE, GREEN] }));
    expect(lastCall(onCommit).brand).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("reads typed text on Enter and on change, as a paste would", () => {
    render(DEFAULT_SETTINGS);
    setValue(pasteField(), "#B27E54, #282623");
    pasteField().dispatchEvent(
      new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
    );
    expect(lastCall(onCommit).brand).toEqual([GOLD, MARBLE]);
    expect(pasteField().value).toBe("");

    setValue(pasteField(), "#111111 #222222 #333333");
    fire(pasteField(), "change");
    expect(lastCall(onCommit).brand).toEqual(["#111111", "#222222", "#333333"]);
    expect(onCommit).toHaveBeenCalledTimes(2);
  });

  it("lets the browser paste when it hands over no clipboard text", () => {
    render();
    const event = new Event("paste", { bubbles: true, cancelable: true });
    pasteField().dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
    expect(onCommit).not.toHaveBeenCalled();
  });
});

describe("drawing again", () => {
  it("follows the settings: rows added and removed in place", () => {
    render();
    render(DEFAULT_SETTINGS);
    expect(rows()).toHaveLength(0);
    render(ten);
    expect(rows()).toHaveLength(8);
    expect(hexField(7).value).toBe(ten.brand[9]);
  });

  it("keeps the controls the user is typing in", () => {
    render();
    hexField(0).focus();
    setValue(hexField(0), "#12");
    render(palette(GOLD, MARBLE, "#445566", SLATE, NEON));
    expect(hexField(0).value).toBe("#12");
    expect(hexField(1).value).toBe(SLATE);
    expect(picker(0).value).toBe("#445566");
  });

  it("commits once per change however often it was drawn", () => {
    render();
    render();
    render();
    setValue(picker(0), "#abcdef");
    fire(picker(0), "change");
    expect(onCommit).toHaveBeenCalledTimes(1);
  });

  it("edits the newest settings after a redraw", () => {
    render();
    render(palette(GOLD, MARBLE, "#445566", SLATE));
    setValue(picker(1), "#abcdef");
    fire(picker(1), "change");
    expect(lastCall(onCommit).brand).toEqual([
      GOLD,
      MARBLE,
      "#445566",
      "#ABCDEF",
    ]);
  });

  it("builds into the container's own document", () => {
    const parsed = new DOMParser().parseFromString(
      '<div id="editor"></div>',
      "text/html",
    );
    const host = parsed.getElementById("editor")!;
    renderPaletteEditor(host, {
      settings: five,
      onPreview: () => undefined,
      onCommit: () => undefined,
    });
    expect(host.querySelectorAll("[data-brand-index]")).toHaveLength(3);
    expect(host.querySelector("#add-brand-color")?.ownerDocument).toBe(parsed);
  });
});

describe("colors that arrive from outside the field", () => {
  it("replaces the list for two or more, as a paste does, and answers the sentence", () => {
    render(DEFAULT_SETTINGS);
    const said = applyBrandColors(container, [GOLD, MARBLE, GREEN]);
    expect(lastCall(onCommit).brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(said).toBe("3 colors read.");
    expect(status().textContent).toBe(said);
  });

  it("appends one color, and refuses it politely at ten", () => {
    render();
    expect(applyBrandColors(container, ["#123456"])).toBe("Added #123456.");
    expect(lastCall(onCommit).brand).toEqual([...five.brand, "#123456"]);

    onCommit.mockClear();
    render(ten);
    expect(applyBrandColors(container, ["#123456"])).toBe(
      "Ten colors is the most, so #123456 was not added. Remove one first.",
    );
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("says its own sentence for an empty list, and changes nothing", () => {
    render();
    expect(
      applyBrandColors(container, [], 0, "No fills in the selection."),
    ).toBe("No fills in the selection.");
    expect(status().textContent).toBe("No fills in the selection.");
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("caps at ten and counts what it left out, with what the caller already dropped", () => {
    render();
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    expect(applyBrandColors(container, twelve, 3)).toBe(
      "10 colors read, 5 more did not fit.",
    );
    expect(lastCall(onCommit).brand).toHaveLength(MAX_BRAND_COLORS);
  });

  it("does nothing for a container that holds no editor", () => {
    const bare = document.createElement("div");
    expect(applyBrandColors(bare, [GOLD, MARBLE], 0, "Nothing here.")).toBe(
      "Nothing here.",
    );
    expect(onCommit).not.toHaveBeenCalled();
  });
});
