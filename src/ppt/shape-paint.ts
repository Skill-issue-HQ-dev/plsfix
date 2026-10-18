// One-click brand colors on the selected shapes: a fill, an outline or a text
// color, or no fill and no outline. The Object tools' pattern (object-tools.ts):
// the PowerPointApi 1.5 guard, getSelectedShapes, one load sync and one write
// sync, both under withSyncDeadline. Invariant: a shape type that refuses is
// skipped and counted, never fatal; a refused batch is told shape by shape.

import { normalizeHex } from "../settings";
import { isDrawTimeout, withSyncDeadline } from "./chart-draw";
import { requireObjectToolsApi } from "./object-tools";

export type PaintMode = "fill" | "outline" | "text";

const STAGE = "Brand color";
const MODES: readonly string[] = ["fill", "outline", "text"];
const LABEL: Record<PaintMode, string> = {
  fill: "Fill",
  outline: "Outline",
  text: "Text",
};
const NONE: Record<PaintMode, string> = {
  fill: "No fill",
  outline: "No outline",
  text: "No text color",
};
const NO_SELECTION = "Select at least one object to color.";

// What each mode may be written to, by the type the host reports. A picture
// has no fill to set, a line no fill and no text, a table and a group take
// their colors cell by cell and child by child: those are skipped, not tried.
const SHAPES = [
  "GeometricShape",
  "TextBox",
  "Placeholder",
  "Callout",
  "Freeform",
];
const TAKES: Record<PaintMode, ReadonlySet<string>> = {
  fill: new Set(SHAPES),
  outline: new Set([...SHAPES, "Line", "Image"]),
  text: new Set(SHAPES),
};

interface Spec {
  mode: PaintMode;
  /** #RRGGBB, or null for None (fill and outline only). */
  color: string | null;
}

// What one write round trip did, or why it was refused.
interface Outcome {
  /** The ids that were asked to take the color. */
  ids: string[];
  /** Selected shapes whose type cannot take it. */
  skipped: number;
  /** The color went to selected text, not to shapes. */
  inText: boolean;
  /** The host's refusal of the batch, when it refused. */
  refusal?: unknown;
}

function isMode(mode: string): mode is PaintMode {
  return MODES.includes(mode);
}

function specOf(mode: string, hex: string | null): Spec {
  if (!isMode(mode)) throw new Error(`${STAGE}: unknown swatch ${mode}.`);
  const color = hex === null ? null : normalizeHex(hex);
  if (hex !== null && color === null) {
    throw new Error(`${STAGE}: ${hex} is not a color.`);
  }
  if (mode === "text" && color === null) {
    throw new Error(`${STAGE}: a text color needs a color.`);
  }
  return { mode, color };
}

function named(spec: Spec): string {
  return spec.color === null
    ? NONE[spec.mode]
    : `${LABEL[spec.mode]} ${spec.color}`;
}

function receipt(
  spec: Spec,
  painted: number,
  skipped: number,
  inText: boolean,
) {
  const where = inText
    ? "on the selected text"
    : `on ${String(painted)} ${painted === 1 ? "object" : "objects"}`;
  const left = skipped > 0 ? `, ${String(skipped)} skipped` : "";
  return `${named(spec)} ${where}${left}.`;
}

function write(shape: PowerPoint.Shape, { mode, color }: Spec): void {
  if (mode === "fill") {
    if (color === null) shape.fill.clear();
    else shape.fill.setSolidColor(color);
  } else if (mode === "outline") {
    shape.lineFormat.visible = color !== null;
    if (color !== null) shape.lineFormat.color = color;
  } else {
    shape.textFrame.textRange.font.color = color!;
  }
}

// The selected text, when some is highlighted: a caret alone is no selection.
function loadTextRange(
  context: PowerPoint.RequestContext,
  spec: Spec,
): PowerPoint.TextRange | null {
  if (spec.mode !== "text") return null;
  const range = context.presentation.getSelectedTextRangeOrNullObject();
  range.load("isNullObject,length");
  return range;
}

function highlighted(range: PowerPoint.TextRange | null): boolean {
  return range !== null && !range.isNullObject && range.length > 0;
}

function nothingTakesIt(spec: Spec, count: number): Error {
  const noun = count === 1 ? "object" : "objects";
  const word = spec.color === null ? NONE[spec.mode] : LABEL[spec.mode];
  return new Error(`${word} does not apply to the selected ${noun}.`);
}

// One load sync, one write sync. A batch the host refuses comes back as an
// outcome carrying the refusal, so the caller can try again shape by shape in a
// run of its own; a batch that never answers is the deadline's, and thrown.
async function batch(
  context: PowerPoint.RequestContext,
  spec: Spec,
): Promise<Outcome> {
  const selected = context.presentation.getSelectedShapes();
  selected.load("items/id,items/type");
  const range = loadTextRange(context, spec);
  await withSyncDeadline(context.sync(), "reading the selection");
  if (range !== null && highlighted(range)) {
    range.font.color = spec.color!;
    await withSyncDeadline(context.sync(), "coloring the text");
    return { ids: [], skipped: 0, inText: true };
  }
  const shapes = selected.items;
  if (shapes.length === 0) throw new Error(NO_SELECTION);
  const takes = TAKES[spec.mode];
  const fit = shapes.filter((shape) => takes.has(String(shape.type)));
  if (fit.length === 0) throw nothingTakesIt(spec, shapes.length);
  for (const shape of fit) write(shape, spec);
  const outcome = {
    ids: fit.map((shape) => shape.id),
    skipped: shapes.length - fit.length,
    inText: false,
  };
  try {
    await withSyncDeadline(context.sync(), "coloring the shapes");
  } catch (error) {
    if (isDrawTimeout(error)) throw error;
    return { ...outcome, refusal: error };
  }
  return outcome;
}

// The batch was refused and is not atomic: what it did stays done, so each
// shape is asked again on its own and the ones the host still refuses are
// skipped and counted. When none takes it the answer is the one for a selection
// nothing takes, never the host's raw text.
async function oneByOne(spec: Spec, outcome: Outcome): Promise<string> {
  return PowerPoint.run(async (context) => {
    const selection = context.presentation.getSelectedShapes();
    let painted = 0;
    for (const id of outcome.ids) {
      try {
        write(selection.getItem(id), spec);
        await withSyncDeadline(context.sync(), "coloring a shape");
        painted += 1;
      } catch (error) {
        if (isDrawTimeout(error)) throw error;
      }
    }
    if (painted === 0) {
      throw nothingTakesIt(spec, outcome.ids.length + outcome.skipped);
    }
    const skipped = outcome.skipped + outcome.ids.length - painted;
    return receipt(spec, painted, skipped, false);
  });
}

/**
 * Paints `hex` onto the selection as its fill, its outline or its text color
 * (`mode` "fill", "outline" or "text"); a null color clears the fill or hides
 * the outline. A Text swatch with text highlighted colors that text only.
 * Answers the receipt the pane toasts: "Fill #B27E54 on 3 objects."
 */
export async function paintShapes(
  mode: string,
  hex: string | null,
): Promise<string> {
  const spec = specOf(mode, hex);
  requireObjectToolsApi();
  const outcome = await PowerPoint.run((context) => batch(context, spec));
  if (outcome.refusal !== undefined) return oneByOne(spec, outcome);
  const painted = outcome.ids.length;
  return receipt(spec, painted, outcome.skipped, outcome.inText);
}
