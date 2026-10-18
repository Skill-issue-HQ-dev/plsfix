// Writes a custom style record (src/custom-styles.ts) onto an Excel cell style:
// the properties of each category the style includes, then the include flags.
// Protection is never included. Invariant: nothing is read back and rewritten;
// every field is built from the record, and a None edge is a style alone.

import type {
  CustomStyle,
  StyleEdge,
  StyleEdgeName,
  StyleFont,
} from "../custom-styles";

// Read inside the call, never at module scope: office.js defines the enums, and
// the pane also loads in a plain browser where they do not exist yet.
function edgeIndexes(): [StyleEdgeName, Excel.BorderIndex][] {
  return [
    ["top", Excel.BorderIndex.edgeTop],
    ["bottom", Excel.BorderIndex.edgeBottom],
    ["left", Excel.BorderIndex.edgeLeft],
    ["right", Excel.BorderIndex.edgeRight],
  ];
}

// An edge the record does not list is no line: a new style must not inherit one.
// A None edge is written as its style alone, on every host: a weight beside it
// can draw the line (the web does, the rig 27.09), and a weight beside Double
// can turn it into a single one, so a double rule keeps the weight Excel gave it.
function writeEdge(
  border: Excel.RangeBorder,
  edge: StyleEdge | undefined,
): void {
  if (!edge || edge.style === Excel.BorderLineStyle.none) {
    border.style = Excel.BorderLineStyle.none;
    return;
  }
  border.style = edge.style as Excel.BorderLineStyle;
  if (
    edge.weight !== undefined &&
    edge.style !== Excel.BorderLineStyle.double
  ) {
    border.weight = edge.weight as Excel.BorderWeight;
  }
  if (edge.color) border.color = edge.color;
}

function writeBorders(style: Excel.Style, spec: CustomStyle): void {
  for (const [name, index] of edgeIndexes()) {
    writeEdge(style.borders.getItem(index), spec.borders?.[name]);
  }
}

function writeFont(style: Excel.Style, font: StyleFont): void {
  if (font.name !== undefined) style.font.name = font.name;
  if (font.size !== undefined) style.font.size = font.size;
  if (font.bold !== undefined) style.font.bold = font.bold;
  if (font.italic !== undefined) style.font.italic = font.italic;
  if (font.color !== undefined) style.font.color = font.color;
}

// Null is no fill: cleared, never a white one.
function writeFill(style: Excel.Style, fill: string | null | undefined): void {
  if (fill === undefined) return;
  if (fill === null) style.fill.clear();
  else style.fill.color = fill;
}

export function writeStyle(style: Excel.Style, spec: CustomStyle): void {
  const { include } = spec;
  if (include.number && spec.numberFormat !== undefined) {
    style.numberFormat = spec.numberFormat;
  }
  if (include.font && spec.font) writeFont(style, spec.font);
  if (include.patterns) writeFill(style, spec.fill);
  if (include.border) writeBorders(style, spec);
  if (include.alignment && spec.horizontalAlignment !== undefined) {
    style.horizontalAlignment =
      spec.horizontalAlignment as Excel.HorizontalAlignment;
  }
  style.includeNumber = include.number;
  style.includeFont = include.font;
  style.includePatterns = include.patterns;
  style.includeBorder = include.border;
  style.includeAlignment = include.alignment;
  style.includeProtection = false;
}
