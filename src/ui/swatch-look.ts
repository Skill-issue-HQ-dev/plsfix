// How one color swatch is drawn, by the mode of its strip: Fill a solid block, Font
// and Text a bold A over a bar of the color (Office's font color button), Outline a
// ring. DOM only; every color is an inline style, the palette being the user's data.
// Invariant: a look draws inside the button; names, titles and tap height stay elsewhere.

export type SwatchLook = "block" | "letter" | "ring";

const LOOKS: Readonly<Record<string, SwatchLook>> = {
  font: "letter",
  text: "letter",
  outline: "ring",
};
const GLYPH = "A";

/** The look a strip's mode draws its colors in: a block unless a mode says otherwise. */
export function lookOf(mode: string): SwatchLook {
  return LOOKS[mode] ?? "block";
}

// A decorative part of a swatch: the button's aria-label is its name.
function part(doc: Document, className: string, text?: string): HTMLElement {
  const node = doc.createElement("span");
  node.className = className;
  node.setAttribute("aria-hidden", "true");
  if (text !== undefined) node.textContent = text;
  return node;
}

/** Paints `hex` onto a color swatch button in the look of `mode`. */
export function applyLook(
  button: HTMLButtonElement,
  mode: string,
  hex: string,
): void {
  const doc = button.ownerDocument;
  switch (lookOf(mode)) {
    case "letter": {
      const glyph = part(doc, "swatch-glyph", GLYPH);
      glyph.style.color = hex;
      const bar = part(doc, "swatch-bar");
      bar.style.background = hex;
      button.classList.add("swatch-letter");
      button.append(glyph, bar);
      break;
    }
    case "ring": {
      const hoop = part(doc, "swatch-hoop");
      hoop.style.borderColor = hex;
      button.classList.add("swatch-ring");
      button.append(hoop);
      break;
    }
    default:
      button.style.background = hex;
  }
}
