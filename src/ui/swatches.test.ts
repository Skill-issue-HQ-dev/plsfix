// @vitest-environment jsdom
// renderSwatches: the rows it builds for a palette, what a press hands to
// onPick (a spy), which swatch is active for the Shades row, and the keyboard
// walk. DOM only.
import {
  afterEach,
  beforeEach,
  describe,
  expect,
  it,
  vi,
  type Mock,
} from "vitest";
import { shadeLadder } from "../palette";
import { renderSwatches, type SwatchOptions } from "./swatches";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const COLORS = [GOLD, MARBLE, GREEN];

const EXCEL_STRIPS: SwatchOptions["strips"] = [
  { mode: "fill", label: "Fill", none: "No fill" },
  { mode: "font", label: "Font" },
];

let container: HTMLElement;
let onPick: Mock<(mode: string, hex: string | null) => void>;

function render(colors: readonly string[] = COLORS): void {
  renderSwatches(container, {
    colors,
    strips: EXCEL_STRIPS,
    onPick: (mode, hex) => onPick(mode, hex),
  });
}

beforeEach(() => {
  document.body.innerHTML = '<div id="host"></div>';
  container = document.getElementById("host")!;
  onPick = vi.fn<(mode: string, hex: string | null) => void>();
});

afterEach(() => {
  vi.useRealTimers();
});

const strip = (mode: string): HTMLElement =>
  container.querySelector<HTMLElement>(`[data-swatch-strip="${mode}"]`)!;
const buttonsOf = (row: HTMLElement): HTMLButtonElement[] => [
  ...row.querySelectorAll<HTMLButtonElement>("button"),
];
const shadesRow = (): HTMLElement =>
  container.querySelector<HTMLElement>("[data-swatch-shades]")!;
const base = (mode: string, hex: string): HTMLButtonElement =>
  strip(mode).querySelector<HTMLButtonElement>(`[data-swatch-hex="${hex}"]`)!;
const caption = (): string =>
  container.querySelector(".swatch-caption")?.textContent ?? "";

function hover(button: HTMLElement, from: HTMLElement | null = null): void {
  button.dispatchEvent(
    new MouseEvent("mouseover", { bubbles: true, relatedTarget: from }),
  );
}
function focus(button: HTMLElement): void {
  button.focus();
}
function press(button: HTMLElement, key: string): KeyboardEvent {
  const event = new KeyboardEvent("keydown", {
    key,
    bubbles: true,
    cancelable: true,
  });
  button.dispatchEvent(event);
  return event;
}

describe("the strips", () => {
  it("builds one labelled strip per mode, in the order given", () => {
    render();
    const rows = [...container.querySelectorAll<HTMLElement>(".swatch-row")];
    expect(rows.map((row) => row.dataset.swatchStrip ?? "shades")).toEqual([
      "fill",
      "font",
      "shades",
    ]);
    expect(strip("fill").querySelector(".swatch-label")?.textContent).toBe(
      "Fill",
    );
    expect(strip("font").getAttribute("role")).toBe("group");
    expect(strip("font").getAttribute("aria-label")).toBe("Font");
  });

  // Every swatch is painted inline, on the button (Fill's block) or on a part
  // of it (Font's A and bar): the looks themselves are swatches.looks.test.ts.
  it("gives every brand color a button in each strip, painted inline", () => {
    render();
    for (const mode of ["fill", "font"]) {
      const swatches = buttonsOf(strip(mode)).filter(
        (button) => button.dataset.swatchHex !== undefined,
      );
      expect(swatches.map((button) => button.dataset.swatchHex)).toEqual(
        COLORS,
      );
      for (const button of swatches) {
        expect(button.type).toBe("button");
        expect(button.dataset.swatchMode).toBe(mode);
        const painted = [button, ...button.querySelectorAll<HTMLElement>("*")];
        expect(
          painted.some(
            (node) => node.style.background !== "" || node.style.color !== "",
          ),
          button.getAttribute("aria-label") ?? "",
        ).toBe(true);
        expect(button.classList.contains("swatch")).toBe(true);
      }
    }
  });

  it("names each button for what a press does, in aria-label and title", () => {
    render();
    expect(base("fill", GOLD).getAttribute("aria-label")).toBe(`Fill ${GOLD}`);
    expect(base("fill", GOLD).title).toBe(`Fill ${GOLD}`);
    expect(base("font", MARBLE).getAttribute("aria-label")).toBe(
      `Font ${MARBLE}`,
    );
    expect(base("font", MARBLE).title).toBe(`Font ${MARBLE}`);
  });

  it("ends the Fill strip on a No fill button, and the Font strip on a color", () => {
    render();
    const none = buttonsOf(strip("fill")).at(-1)!;
    expect(none.getAttribute("aria-label")).toBe("No fill");
    expect(none.title).toBe("No fill");
    expect(none.dataset.swatchMode).toBe("fill");
    expect(none.dataset.swatchHex).toBeUndefined();
    expect(buttonsOf(strip("font")).at(-1)!.dataset.swatchHex).toBe(GREEN);
    expect(
      buttonsOf(strip("font")).some((button) =>
        button.classList.contains("swatch-none"),
      ),
    ).toBe(false);
  });

  it("takes any strips it is given (PowerPoint passes three, none with No outline)", () => {
    renderSwatches(container, {
      colors: COLORS,
      strips: [
        { mode: "fill", label: "Fill", none: "No fill" },
        { mode: "outline", label: "Outline", none: "No outline" },
        { mode: "text", label: "Text" },
      ],
      onPick: (mode, hex) => onPick(mode, hex),
    });
    expect(
      [...container.querySelectorAll("[data-swatch-strip]")].map(
        (row) => (row as HTMLElement).dataset.swatchStrip,
      ),
    ).toEqual(["fill", "outline", "text"]);
    expect(
      container
        .querySelector('[data-swatch-strip="outline"] .swatch-none')
        ?.getAttribute("aria-label"),
    ).toBe("No outline");
  });

  it("skips a color that is not #RRGGBB instead of painting it", () => {
    render([GOLD, "not a color", MARBLE]);
    expect(
      buttonsOf(strip("fill"))
        .map((button) => button.dataset.swatchHex)
        .filter(Boolean),
    ).toEqual([GOLD, MARBLE]);
  });
});

describe("the Shades row", () => {
  it("starts on the first color of the first strip and says so", () => {
    render();
    expect(caption()).toBe(`Fill shades of ${GOLD}`);
    expect(shadesRow().getAttribute("aria-label")).toBe(
      `Fill shades of ${GOLD}`,
    );
    expect(base("fill", GOLD).getAttribute("aria-current")).toBe("true");
    expect(base("font", GOLD).hasAttribute("aria-current")).toBe(false);
  });

  it("holds the five shadeLadder buttons of the active color, in the active mode", () => {
    render();
    const shades = buttonsOf(shadesRow());
    const ladder = shadeLadder(GOLD);
    expect(shades).toHaveLength(5);
    shades.forEach((button, at) => {
      const { hex, label } = ladder[at]!;
      expect(button.dataset.swatchMode).toBe("fill");
      expect(button.dataset.swatchHex).toBe(hex);
      expect(button.getAttribute("aria-label")).toBe(`Fill ${hex}, ${label}`);
      expect(button.title).toBe(`Fill ${hex}, ${label}`);
      expect(button.style.background).not.toBe("");
    });
  });

  it("follows the swatch the pointer rests on, and applies in that strip's mode", () => {
    render();
    hover(base("font", MARBLE));
    expect(caption()).toBe(`Font shades of ${MARBLE}`);
    const shades = buttonsOf(shadesRow());
    expect(shades.every((button) => button.dataset.swatchMode === "font")).toBe(
      true,
    );
    expect(shades[0]!.dataset.swatchHex).toBe(shadeLadder(MARBLE)[0]!.hex);
    expect(base("font", MARBLE).getAttribute("aria-current")).toBe("true");
    expect(base("fill", GOLD).hasAttribute("aria-current")).toBe(false);
  });

  it("follows keyboard focus on a base swatch", () => {
    render();
    focus(base("fill", GREEN));
    expect(caption()).toBe(`Fill shades of ${GREEN}`);
  });

  it("is not moved by focus or a pointer on a shade, the none button or a label", () => {
    render();
    hover(base("fill", MARBLE));
    focus(buttonsOf(shadesRow())[1]!);
    hover(buttonsOf(shadesRow())[2]!);
    hover(buttonsOf(strip("fill")).at(-1)!);
    hover(strip("font").querySelector(".swatch-label")!);
    expect(caption()).toBe(`Fill shades of ${MARBLE}`);
  });

  it("settles on the swatch last pressed when nothing was hovered or focused", () => {
    render();
    base("font", GREEN).click();
    expect(caption()).toBe(`Font shades of ${GREEN}`);
  });

  it("keeps the active swatch when the palette is drawn again", () => {
    render();
    hover(base("font", GREEN));
    render([GOLD, MARBLE, GREEN, "#41564F"]);
    expect(caption()).toBe(`Font shades of ${GREEN}`);
    expect(base("font", GREEN).getAttribute("aria-current")).toBe("true");
  });

  it("falls back to the first color when the active one left the palette", () => {
    render();
    hover(base("font", GREEN));
    render([MARBLE, GOLD]);
    expect(caption()).toBe(`Fill shades of ${MARBLE}`);
  });

  it("shows no shades while the palette has no colors", () => {
    render([]);
    expect(buttonsOf(shadesRow())).toHaveLength(0);
    expect(caption()).toBe("");
  });
});

describe("a press", () => {
  it("hands the mode and the color to onPick", () => {
    render();
    base("fill", GOLD).click();
    base("font", GREEN).click();
    expect(onPick.mock.calls).toEqual([
      ["fill", GOLD],
      ["font", GREEN],
    ]);
  });

  it("hands null for the No fill button", () => {
    render();
    buttonsOf(strip("fill")).at(-1)!.click();
    expect(onPick).toHaveBeenCalledExactlyOnceWith("fill", null);
  });

  it("hands the shade's own color in the active swatch's mode", () => {
    render();
    hover(base("font", MARBLE));
    const shade = buttonsOf(shadesRow())[3]!;
    shade.click();
    expect(onPick).toHaveBeenCalledExactlyOnceWith(
      "font",
      shadeLadder(MARBLE)[3]!.hex,
    );
  });

  it("ignores a click that is on no swatch", () => {
    render();
    strip("fill").querySelector<HTMLElement>(".swatch-label")!.click();
    container.click();
    expect(onPick).not.toHaveBeenCalled();
  });

  it("runs onPick once however often the palette was drawn: one delegated listener", () => {
    render();
    render();
    render([GOLD, MARBLE]);
    base("fill", GOLD).click();
    expect(onPick).toHaveBeenCalledTimes(1);
  });

  it("calls the newest onPick after a redraw", () => {
    render();
    const newer = vi.fn();
    renderSwatches(container, {
      colors: COLORS,
      strips: EXCEL_STRIPS,
      onPick: newer,
    });
    base("fill", GOLD).click();
    expect(onPick).not.toHaveBeenCalled();
    expect(newer).toHaveBeenCalledExactlyOnceWith("fill", GOLD);
  });
});

describe("drawing again", () => {
  it("replaces the buttons instead of piling up rows", () => {
    render();
    render();
    expect(container.querySelectorAll(".swatch-row")).toHaveLength(3);
    expect(strip("fill").querySelectorAll("button")).toHaveLength(
      COLORS.length + 1,
    );
  });

  it("shows a changed palette", () => {
    render();
    render([GOLD, MARBLE]);
    expect(
      buttonsOf(strip("font")).map((button) => button.dataset.swatchHex),
    ).toEqual([GOLD, MARBLE]);
  });

  it("keeps new buttons disabled while the pane is busy", () => {
    render();
    for (const button of container.querySelectorAll("button")) {
      button.disabled = true;
    }
    render([GOLD, MARBLE, GREEN, "#41564F"]);
    expect(
      [...container.querySelectorAll("button")].every(
        (button) => button.disabled,
      ),
    ).toBe(true);
  });

  it("builds into the container's own document, so a parsed pane gets the same buttons", () => {
    const parsed = new DOMParser().parseFromString(
      '<div id="host"></div>',
      "text/html",
    );
    const host = parsed.getElementById("host")!;
    renderSwatches(host, {
      colors: COLORS,
      strips: EXCEL_STRIPS,
      onPick: () => undefined,
    });
    expect(host.querySelectorAll("button[data-swatch-mode]").length).toBe(
      COLORS.length * 2 + 1 + 5,
    );
    expect(host.firstElementChild?.ownerDocument).toBe(parsed);
  });
});

describe("a pointer sweeping across the other strip", () => {
  beforeEach(() => {
    vi.useFakeTimers();
  });

  it("does not retarget the shades on its way down to them", () => {
    render();
    hover(base("fill", GOLD));
    vi.advanceTimersByTime(40);
    // Straight down: the Font strip's swatch under the pointer, a moment later.
    hover(base("font", GOLD), strip("fill"));
    expect(caption()).toBe(`Fill shades of ${GOLD}`);
    vi.advanceTimersByTime(40);
    hover(buttonsOf(shadesRow())[0]!, strip("font"));
    vi.advanceTimersByTime(500);
    expect(caption()).toBe(`Fill shades of ${GOLD}`);
  });

  it("retargets once the pointer rests on the other strip", () => {
    render();
    hover(base("fill", GOLD));
    vi.advanceTimersByTime(40);
    hover(base("font", MARBLE), strip("fill"));
    vi.advanceTimersByTime(200);
    expect(caption()).toBe(`Font shades of ${MARBLE}`);
  });

  it("retargets at once for a swatch in the same strip or after a rest", () => {
    render();
    hover(base("fill", GOLD));
    hover(base("fill", MARBLE));
    expect(caption()).toBe(`Fill shades of ${MARBLE}`);
    vi.advanceTimersByTime(2_000);
    hover(base("font", GREEN));
    expect(caption()).toBe(`Font shades of ${GREEN}`);
  });

  it("drops a pending retarget when the pointer leaves the swatches", () => {
    render();
    hover(base("fill", GOLD));
    vi.advanceTimersByTime(40);
    hover(base("font", MARBLE), strip("fill"));
    container.dispatchEvent(new MouseEvent("mouseleave"));
    vi.advanceTimersByTime(500);
    expect(caption()).toBe(`Fill shades of ${GOLD}`);
  });
});

describe("the keyboard", () => {
  it("makes each strip one Tab stop: a single tabindex 0, the rest -1", () => {
    render();
    for (const row of [strip("fill"), strip("font"), shadesRow()]) {
      const stops = buttonsOf(row).filter((button) => button.tabIndex === 0);
      expect(stops, row.getAttribute("aria-label") ?? "").toHaveLength(1);
      expect(stops[0]).toBe(buttonsOf(row)[0]);
    }
  });

  it("moves along a strip with Left and Right, stopping at its ends", () => {
    render();
    const [first, second, third] = buttonsOf(strip("fill"));
    first!.focus();
    const event = press(first!, "ArrowRight");
    expect(event.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(second);
    expect(second!.tabIndex).toBe(0);
    expect(first!.tabIndex).toBe(-1);
    press(second!, "ArrowRight");
    expect(document.activeElement).toBe(third);
    press(third!, "ArrowLeft");
    expect(document.activeElement).toBe(second);
    press(second!, "ArrowLeft");
    expect(document.activeElement).toBe(first);
    press(first!, "ArrowLeft");
    expect(document.activeElement).toBe(first);
    const last = buttonsOf(strip("fill")).at(-1)!;
    last.focus();
    press(last, "ArrowRight");
    expect(document.activeElement).toBe(last);
  });

  it("jumps to the ends with Home and End", () => {
    render();
    const all = buttonsOf(strip("font"));
    all[1]!.focus();
    press(all[1]!, "End");
    expect(document.activeElement).toBe(all.at(-1));
    press(all.at(-1)!, "Home");
    expect(document.activeElement).toBe(all[0]);
  });

  it("goes Down from a base swatch to the Shades row with that color active, and Up returns", () => {
    render();
    const marble = base("font", MARBLE);
    marble.focus();
    const down = press(marble, "ArrowDown");
    expect(down.defaultPrevented).toBe(true);
    expect(caption()).toBe(`Font shades of ${MARBLE}`);
    const shades = buttonsOf(shadesRow());
    expect(document.activeElement).toBe(shades[0]);
    expect(shades[0]!.tabIndex).toBe(0);

    press(shades[0]!, "ArrowRight");
    expect(document.activeElement).toBe(shades[1]);
    const up = press(shades[1]!, "ArrowUp");
    expect(up.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(base("font", MARBLE));
  });

  it("does nothing on Down from the No fill button", () => {
    render();
    const none = buttonsOf(strip("fill")).at(-1)!;
    none.focus();
    const event = press(none, "ArrowDown");
    expect(event.defaultPrevented).toBe(false);
    expect(document.activeElement).toBe(none);
  });

  it("leaves Enter, Space and Tab to the browser", () => {
    render();
    const gold = base("fill", GOLD);
    for (const key of ["Enter", " ", "Tab"]) {
      expect(press(gold, key).defaultPrevented, key).toBe(false);
    }
  });

  it("keeps the Tab stop where focus last was", () => {
    render();
    const marble = base("fill", MARBLE);
    marble.focus();
    expect(marble.tabIndex).toBe(0);
    expect(base("fill", GOLD).tabIndex).toBe(-1);
  });
});
