// A shape's look in and out of PowerPoint, for Smart Painter and the named styles alike:
// capture one selected shape (fill, outline, text look), paint a look onto every selected
// shape, each sync under withSyncDeadline. Invariant: a shape that refuses is skipped and
// counted, never fatal, and a value the host could not say is never written back.

import { readLook } from "../shape-styles-codec";
import type { ShapeLook } from "../shape-styles";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import { requireObjectToolsApi } from "./shapes";

type Part = "fill" | "outline" | "text";

// The types that carry a text frame: PowerPoint refuses the frame of a picture, a
// table, a group and a line (src/ppt/placement.ts measured it), so a text look is
// neither read from nor written to those. The types shape-paint.ts's text mode takes.
const TEXT_TYPES: ReadonlySet<string> = new Set([
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Callout",
  "Freeform",
]);

function requireCount(
  count: number,
  minimum: number,
  message: string,
  maximum = Number.POSITIVE_INFINITY,
): void {
  if (count < minimum || count > maximum) throw new Error(message);
}

// ---------------------------------------------------------------------------
// Capture
// ---------------------------------------------------------------------------

/** What the caller says when the selection or the fill is not one it can take. */
export interface CaptureWords {
  selectOne: string;
  notSolid: string;
}

function loadLook(shape: PowerPoint.Shape, withText: boolean): void {
  shape.fill.load("type,foregroundColor,transparency");
  shape.lineFormat.load("visible,color,transparency,weight,dashStyle,style");
  if (!withText) return;
  shape.textFrame.load("hasText");
  shape.textFrame.textRange.font.load("name,size,color,bold,italic");
}

const FILL_UNREADABLE = "The fill color of this object could not be read.";

function lookOf(
  shape: PowerPoint.Shape,
  withText: boolean,
  words: CaptureWords,
): ShapeLook {
  const type = String(shape.fill.type);
  // The whitelist refuses every other ShapeFillType uniformly - Gradient, Pattern,
  // PictureAndTexture, SlideBackground - so a picture and a gradient meet the same
  // refusal, not two different code paths.
  if (type !== "Solid" && type !== "NoFill") throw new Error(words.notSolid);
  const solid = type === "Solid";
  const line = shape.lineFormat;
  const hasText = withText && shape.textFrame.hasText;
  const range = hasText ? shape.textFrame.textRange.font : undefined;
  const look = readLook({
    fill: solid ? shape.fill.foregroundColor : null,
    fillTransparency: solid ? shape.fill.transparency : undefined,
    // The host answers null, not the string "null", for what a hidden line or a
    // line still on the theme's own look has no value for: readLook leaves those out.
    line: {
      visible: line.visible,
      color: line.color,
      weight: line.weight,
      dashStyle: line.dashStyle,
      transparency: line.transparency,
      style: line.style,
    },
    text: range && {
      name: range.name,
      size: range.size,
      color: range.color,
      bold: range.bold,
      italic: range.italic,
    },
  });
  // A solid fill whose color the host could not say is not a copy to build on: a
  // look saved without its fill would paint the next shape with none of it.
  if (solid && look.fill === undefined) throw new Error(FILL_UNREADABLE);
  return look;
}

/**
 * The look of the one selected shape: its solid or no fill, its outline and,
 * when it holds text, its font name, size, color, bold and italic. Two round
 * trips: the selection, then the look.
 */
export async function captureShapeLook(
  words: CaptureWords,
): Promise<ShapeLook> {
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load("items/id,items/type");
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, 1, words.selectOne, 1);
    const shape = selected.items[0]!;
    const withText = TEXT_TYPES.has(String(shape.type));
    loadLook(shape, withText);
    await withSyncDeadline(context.sync(), "reading the style");
    return lookOf(shape, withText, words);
  });
}

// ---------------------------------------------------------------------------
// Paint
// ---------------------------------------------------------------------------

/** What the caller says when nothing is selected. */
export interface PaintWords {
  selectSome: string;
}

/** The shapes the look went onto, and the selected ones it did not. */
export interface PaintReport {
  painted: number;
  skipped: number;
}

// The parts of the look one shape is asked to take: what the look names, and the
// text only where there is a text frame.
function partsOf(look: ShapeLook, type: string): Part[] {
  const parts: Part[] = [];
  if (look.fill !== undefined) parts.push("fill");
  if (look.line !== undefined) parts.push("outline");
  if (look.text !== undefined && TEXT_TYPES.has(type)) parts.push("text");
  return parts;
}

function writeFill(shape: PowerPoint.Shape, look: ShapeLook): void {
  if (look.fill === undefined) return;
  if (look.fill === null) {
    shape.fill.clear();
    return;
  }
  shape.fill.setSolidColor(look.fill);
  if (look.fillTransparency !== undefined) {
    shape.fill.transparency = look.fillTransparency;
  }
}

// An invisible line has no color, weight, dash or style: only a visible line's
// details are written, and only the ones the look holds.
function writeOutline(shape: PowerPoint.Shape, look: ShapeLook): void {
  const line = look.line!;
  shape.lineFormat.visible = line.visible;
  if (!line.visible) return;
  if (line.color !== undefined) shape.lineFormat.color = line.color;
  if (line.transparency !== undefined) {
    shape.lineFormat.transparency = line.transparency;
  }
  if (line.weight !== undefined) shape.lineFormat.weight = line.weight;
  if (line.dashStyle !== undefined) {
    shape.lineFormat.dashStyle =
      line.dashStyle as PowerPoint.ShapeLineDashStyle;
  }
  if (line.style !== undefined) {
    shape.lineFormat.style = line.style as PowerPoint.ShapeLineStyle;
  }
}

function writeText(shape: PowerPoint.Shape, look: ShapeLook): void {
  const text = look.text!;
  const font = shape.textFrame.textRange.font;
  if (text.name !== undefined) font.name = text.name;
  if (text.size !== undefined) font.size = text.size;
  if (text.color !== undefined) font.color = text.color;
  if (text.bold !== undefined) font.bold = text.bold;
  if (text.italic !== undefined) font.italic = text.italic;
}

const WRITE: Record<Part, (shape: PowerPoint.Shape, look: ShapeLook) => void> =
  { fill: writeFill, outline: writeOutline, text: writeText };

interface Ask {
  id: string;
  parts: Part[];
}

// What one write round trip did, or why it was refused.
interface Outcome {
  /** Selected shapes. */
  total: number;
  /** The shapes the look has something for, and what to ask each. */
  asks: Ask[];
  /** The host's refusal of the batch, when it refused. */
  refusal?: unknown;
}

// One load sync, one write sync (none when the look has nothing for any shape).
// A batch the host refuses comes back as an outcome carrying the refusal, so the
// caller can ask again shape by shape in a run of its own; a batch that never
// answers is the deadline's, and thrown.
async function batch(
  context: PowerPoint.RequestContext,
  look: ShapeLook,
  words: PaintWords,
): Promise<Outcome> {
  const selected = context.presentation.getSelectedShapes();
  selected.load("items/id,items/type");
  await withSyncDeadline(context.sync(), "reading the selection");
  requireCount(selected.items.length, 1, words.selectSome);
  const asks: Ask[] = [];
  for (const shape of selected.items) {
    const parts = partsOf(look, String(shape.type));
    if (parts.length === 0) continue;
    for (const part of parts) WRITE[part](shape, look);
    asks.push({ id: shape.id, parts });
  }
  const outcome: Outcome = { total: selected.items.length, asks };
  if (asks.length === 0) return outcome;
  try {
    await withSyncDeadline(context.sync(), "painting the style");
  } catch (error) {
    if (isDrawTimeout(error)) throw error;
    return { ...outcome, refusal: error };
  }
  return outcome;
}

// Whether the shape took any part: all of them in one round trip, and when the
// host refuses that, each part on its own (a refusal halts a batch where it
// stands, so what came before it landed and what came after never ran).
async function askShape(
  context: PowerPoint.RequestContext,
  selection: PowerPoint.ShapeScopedCollection,
  look: ShapeLook,
  ask: Ask,
): Promise<boolean> {
  const write = async (parts: Part[]): Promise<boolean> => {
    try {
      for (const part of parts) WRITE[part](selection.getItem(ask.id), look);
      await withSyncDeadline(context.sync(), "painting a shape");
      return true;
    } catch (error) {
      if (isDrawTimeout(error)) throw error;
      return false;
    }
  };
  if (await write(ask.parts)) return true;
  if (ask.parts.length === 1) return false;
  let took = false;
  for (const part of ask.parts) took = (await write([part])) || took;
  return took;
}

// The batch was refused and is not atomic: what it did stays done, so each shape
// is asked again on its own, and the ones the host still refuses are skipped.
async function oneByOne(look: ShapeLook, asks: Ask[]): Promise<number> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    let painted = 0;
    for (const ask of asks) {
      if (await askShape(context, selection, look, ask)) painted += 1;
    }
    return painted;
  });
}

/**
 * Puts the look on every selected shape. A part the look does not name is left
 * as it is; the text goes only to shapes that carry a text frame. A shape the
 * host refuses is skipped, never fatal: the report says how many were painted
 * and how many were not, and the caller words it.
 */
export async function paintShapeLook(
  look: ShapeLook,
  words: PaintWords,
): Promise<PaintReport> {
  requireObjectToolsApi();
  const outcome = await PowerPoint.run((context) =>
    batch(context, look, words),
  );
  const painted =
    outcome.refusal === undefined
      ? outcome.asks.length
      : await oneByOne(look, outcome.asks);
  return { painted, skipped: outcome.total - painted };
}
