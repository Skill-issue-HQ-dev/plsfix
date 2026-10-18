// @vitest-environment jsdom
// How each swatch is drawn (src/ui/swatch-look.ts): Fill the solid block, Font and
// Text a bold A over a bar of the color, Outline a ring; the Shades row follows the
// mode of the swatch it shows; labels, titles and the None words do not change with
// the look. DOM only; every color is an inline style.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { shadeLadder } from "../palette";
import { lookOf } from "./swatch-look";
import { renderSwatches, type SwatchStrip } from "./swatches";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const COLORS = [GOLD, MARBLE, GREEN];

const STRIPS: readonly SwatchStrip[] = [
  { mode: "fill", label: "Fill", none: "No fill" },
  { mode: "font", label: "Font" },
  { mode: "outline", label: "Outline", none: "No outline" },
  { mode: "text", label: "Text" },
];

let container: HTMLElement;

beforeEach(() => {
  document.body.innerHTML = '<div id="host"></div>';
  container = document.getElementById("host")!;
  renderSwatches(container, {
    colors: COLORS,
    strips: STRIPS,
    onPick: () => undefined,
  });
});

afterEach(() => {
  vi.useRealTimers();
});

// The way the browser reads a color back out of an inline style.
function inline(hex: string): string {
  const probe = document.createElement("i");
  probe.style.color = hex;
  return probe.style.color;
}

const base = (mode: string, hex: string): HTMLButtonElement =>
  container.querySelector<HTMLButtonElement>(
    `[data-swatch-strip="${mode}"] [data-swatch-hex="${hex}"]`,
  )!;
const shades = (): HTMLButtonElement[] => [
  ...container.querySelectorAll<HTMLButtonElement>(
    "[data-swatch-shades] button",
  ),
];
const part = (button: HTMLElement, selector: string): HTMLElement =>
  button.querySelector<HTMLElement>(selector)!;

function hover(button: HTMLElement): void {
  button.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
}

describe("the look of a mode", () => {
  it("is a ring for Outline, an A over a bar for Font and Text, a block for the rest", () => {
    expect(lookOf("outline")).toBe("ring");
    expect(lookOf("font")).toBe("letter");
    expect(lookOf("text")).toBe("letter");
    expect(lookOf("fill")).toBe("block");
    expect(lookOf("anything else")).toBe("block");
  });
});

describe("Fill: the solid block, as before", () => {
  it("paints the button itself and holds nothing inside", () => {
    for (const hex of COLORS) {
      const button = base("fill", hex);
      expect(button.style.background).toBe(inline(hex));
      expect(button.children).toHaveLength(0);
      expect(button.className).toBe("swatch");
    }
  });
});

describe.each(["font", "text"])(
  "%s: a bold A over a bar of the color",
  (mode) => {
    it("draws the A and the bar in the color, on the pane's own ground", () => {
      for (const hex of COLORS) {
        const button = base(mode, hex);
        expect(button.classList.contains("swatch")).toBe(true);
        expect(button.classList.contains("swatch-letter")).toBe(true);
        // The ground is the sheet's paper: the button carries no color of its own.
        expect(button.style.background).toBe("");
        const glyph = part(button, ".swatch-glyph");
        expect(glyph.textContent).toBe("A");
        expect(glyph.style.color).toBe(inline(hex));
        const bar = part(button, ".swatch-bar");
        expect(bar.textContent).toBe("");
        expect(bar.style.background).toBe(inline(hex));
        expect(button.children).toHaveLength(2);
      }
    });

    it("keeps its name for assistive technology: the parts are decoration", () => {
      const button = base(mode, GOLD);
      expect(part(button, ".swatch-glyph").getAttribute("aria-hidden")).toBe(
        "true",
      );
      expect(part(button, ".swatch-bar").getAttribute("aria-hidden")).toBe(
        "true",
      );
    });
  },
);

describe("Outline: a ring", () => {
  it("draws a thick border in the color around an empty center", () => {
    for (const hex of COLORS) {
      const button = base("outline", hex);
      expect(button.classList.contains("swatch")).toBe(true);
      expect(button.classList.contains("swatch-ring")).toBe(true);
      expect(button.style.background).toBe("");
      const hoop = part(button, ".swatch-hoop");
      expect(hoop.style.borderColor).toBe(inline(hex));
      expect(hoop.style.background).toBe("");
      expect(hoop.textContent).toBe("");
      expect(hoop.children).toHaveLength(0);
      expect(hoop.getAttribute("aria-hidden")).toBe("true");
      expect(button.children).toHaveLength(1);
    }
  });
});

describe("the Shades row follows the mode of the swatch it shows", () => {
  it("is five blocks under a Fill swatch", () => {
    hover(base("fill", GOLD));
    const row = shades();
    expect(row).toHaveLength(5);
    shadeLadder(GOLD).forEach((shade, at) => {
      expect(row[at]!.style.background).toBe(inline(shade.hex));
      expect(row[at]!.children).toHaveLength(0);
    });
  });

  it.each(["font", "text"])("is five colored A's under a %s swatch", (mode) => {
    hover(base(mode, MARBLE));
    const row = shades();
    expect(row).toHaveLength(5);
    shadeLadder(MARBLE).forEach((shade, at) => {
      expect(row[at]!.classList.contains("swatch-letter")).toBe(true);
      expect(part(row[at]!, ".swatch-glyph").style.color).toBe(
        inline(shade.hex),
      );
      expect(part(row[at]!, ".swatch-bar").style.background).toBe(
        inline(shade.hex),
      );
      expect(row[at]!.style.background).toBe("");
    });
  });

  it("is five rings under an Outline swatch", () => {
    hover(base("outline", GREEN));
    const row = shades();
    expect(row).toHaveLength(5);
    shadeLadder(GREEN).forEach((shade, at) => {
      expect(row[at]!.classList.contains("swatch-ring")).toBe(true);
      expect(part(row[at]!, ".swatch-hoop").style.borderColor).toBe(
        inline(shade.hex),
      );
      expect(row[at]!.style.background).toBe("");
    });
  });

  it("changes look as the pointer rests on another mode's swatch", () => {
    // A pointer crossing a strip on its way down is given a moment (swatches.ts).
    vi.useFakeTimers();
    hover(base("outline", GOLD));
    expect(shades().every((each) => each.matches(".swatch-ring"))).toBe(true);
    vi.advanceTimersByTime(500);
    hover(base("font", GOLD));
    expect(shades().every((each) => each.matches(".swatch-letter"))).toBe(true);
    vi.advanceTimersByTime(500);
    hover(base("fill", GOLD));
    expect(shades().every((each) => each.children.length === 0)).toBe(true);
  });
});

describe("what a look leaves alone", () => {
  it("names every button as before: the mode, the hex, the shade's label", () => {
    expect(base("font", GOLD).getAttribute("aria-label")).toBe(`Font ${GOLD}`);
    expect(base("font", GOLD).title).toBe(`Font ${GOLD}`);
    expect(base("text", MARBLE).getAttribute("aria-label")).toBe(
      `Text ${MARBLE}`,
    );
    expect(base("outline", GREEN).getAttribute("aria-label")).toBe(
      `Outline ${GREEN}`,
    );
    expect(base("outline", GREEN).title).toBe(`Outline ${GREEN}`);
    hover(base("outline", GREEN));
    const [first] = shades();
    const { hex, label } = shadeLadder(GREEN)[0]!;
    expect(first!.getAttribute("aria-label")).toBe(`Outline ${hex}, ${label}`);
    expect(first!.title).toBe(`Outline ${hex}, ${label}`);
  });

  it("keeps the None buttons as words, with their own names, in any strip", () => {
    for (const [mode, name] of [
      ["fill", "No fill"],
      ["outline", "No outline"],
    ] as const) {
      const none = container.querySelector<HTMLButtonElement>(
        `[data-swatch-strip="${mode}"] .swatch-none`,
      )!;
      expect(none.textContent).toBe("None");
      expect(none.getAttribute("aria-label")).toBe(name);
      expect(none.title).toBe(name);
      expect(none.children).toHaveLength(0);
      expect(none.className).toBe("swatch swatch-none");
    }
  });

  it("still marks one base swatch as the current one, whatever its look", () => {
    hover(base("outline", MARBLE));
    expect(base("outline", MARBLE).getAttribute("aria-current")).toBe("true");
    expect(base("font", MARBLE).hasAttribute("aria-current")).toBe(false);
  });
});

describe("the sheet", () => {
  const css = (name: string): string =>
    readFileSync(join(process.cwd(), "src/styles", name), "utf8");

  // The rule body of a selector written as `selector {` in a sheet.
  function body(sheet: string, selector: string): string {
    const at = sheet.indexOf(`${selector} {`);
    expect(at, `${selector} has no rule`).toBeGreaterThanOrEqual(0);
    return sheet.slice(sheet.indexOf("{", at) + 1, sheet.indexOf("}", at));
  }

  it("keeps every swatch one 32 px tap target tall, whatever its look", () => {
    expect(css("tokens.css")).toMatch(/--swatch-h:\s*32px/);
    const swatches = css("swatches.css");
    expect(body(swatches, ".swatch")).toMatch(/height:\s*var\(--swatch-h\)/);
    // A look draws inside the button: the letter's rule lays its parts out and
    // sets no height, the ring fills the height the button gives it.
    expect(body(swatches, ".swatch-letter")).not.toMatch(
      /(?:^|[\s;])(?:min-)?height:/,
    );
    expect(body(swatches, ".swatch-hoop")).toMatch(/height:\s*100%/);
  });

  it("draws the A bold through the weight token, at a size the gate reads (11 px up)", () => {
    const glyph = body(css("swatches.css"), ".swatch-glyph");
    expect(glyph).toMatch(/font-weight:\s*var\(--w-strong\)/);
    expect(glyph).toMatch(
      /font-size:\s*var\(--fs-(?:body|title|lead|display)\)/,
    );
  });

  it("gives the ring a thick border whose color the script sets", () => {
    const hoop = body(css("swatches.css"), ".swatch-hoop");
    expect(hoop).toMatch(/border:\s*var\(--swatch-ring-w\) solid/);
    expect(hoop).not.toMatch(/background/);
  });
});
