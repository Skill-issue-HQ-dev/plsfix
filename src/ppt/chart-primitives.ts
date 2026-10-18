// One chart primitive turned into one PowerPoint add: a rectangle, a pie
// slice, an ellipse, a triangle, a block arc, a donut, a label or a line, plus
// the adjustments a slice, a triangle, an arc and a donut are given: in the
// batch that adds them on the web, in the sync after their add on the desktop.
// Split out of chart-draw.ts at its 400-line cap. Invariant: every shape is
// added at the primitive's own box moved to the chart's corner, and nothing
// here syncs.

import type { Line, Primitive, Text } from "../chart-shapes";
import type { Triangle } from "../chart-shapes-types";
import type { Box } from "../layout";
import type { DrawHost } from "./chart-cost";

const ALIGNMENT = { l: "Left", c: "Center", r: "Right" } as const;
// A tip-down triangle is the standing one turned half a turn about the centre
// of its box; the box stays the unrotated box (S1).
const HALF_TURN = 180;

// A shape queued in a batch, with the primitive it draws.
export interface Added {
  shape: PowerPoint.Shape;
  primitive: Primitive;
}

function offset(inner: Box, box: Box): Box {
  return {
    left: box.left + inner.left,
    top: box.top + inner.top,
    width: inner.width,
    height: inner.height,
  };
}

function filled(shape: PowerPoint.Shape, color: string): PowerPoint.Shape {
  shape.fill.setSolidColor(color);
  shape.lineFormat.visible = false;
  return shape;
}

// Six property writes: the font, its size, its colour, the alignment, and -
// since even a box the layout sized for PowerPoint's own insets can still
// wrap the moment the real font runs a hair wider than the estimate - no word
// wrap and no autosize, so the host shows exactly the box's one line instead
// of stacking it ("1 / 519"). The margins are left exactly as the host made
// them.
function addLabel(
  shapes: PowerPoint.ShapeCollection,
  text: Text,
  font: string,
  at: Box,
): PowerPoint.Shape {
  const shape = shapes.addTextBox(text.text, at);
  const range = shape.textFrame.textRange;
  range.font.name = font;
  range.font.size = text.size;
  range.font.color = text.color;
  range.font.bold = text.bold;
  range.paragraphFormat.horizontalAlignment = ALIGNMENT[text.align];
  shape.textFrame.wordWrap = false;
  shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
  return shape;
}

// A box is always normalised to a non-negative width and height (see
// chart-shapes-parts.ts lineBetween), so a rising segment - bottom-left to
// top-right - cannot be drawn as a straight connector from it: the host has
// no "flip" on ConnectorType.straight, only a second preset for the other
// diagonal. Everything else keeps the connector it always drew as.
function addLineShape(
  shapes: PowerPoint.ShapeCollection,
  primitive: Line,
  at: Box,
): PowerPoint.Shape {
  if (primitive.rising) {
    const shape = shapes.addGeometricShape(
      PowerPoint.GeometricShapeType.lineInverse,
      at,
    );
    shape.lineFormat.color = primitive.color;
    shape.lineFormat.weight = primitive.weight;
    return shape;
  }
  const shape = shapes.addLine(PowerPoint.ConnectorType.straight, at);
  // The host reads a zero width or height in the add as "not given" and
  // draws the line sloped over its 72 pt default (PowerPoint for the web,
  // 30.08); written after the add, a zero side sticks.
  shape.width = at.width;
  shape.height = at.height;
  shape.lineFormat.color = primitive.color;
  shape.lineFormat.weight = primitive.weight;
  return shape;
}

// A primitive is one shape: a bar, a legend swatch and a stacked segment are
// rectangles, a slice is a Pie (or an Ellipse when it is the whole circle),
// a baseline or a waterfall connector is a line (straight, or the host's
// other diagonal preset for a rising one). `box` is the chart's corner and
// `font` the brand font every label wears.
export function addPrimitive(
  shapes: PowerPoint.ShapeCollection,
  primitive: Primitive,
  box: Box,
  font: string,
): PowerPoint.Shape {
  const at = offset(primitive.box, box);
  switch (primitive.kind) {
    case "text":
      return addLabel(shapes, primitive, font, at);
    case "line":
      return addLineShape(shapes, primitive, at);
    case "wedge":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.pie, at),
        primitive.color,
      );
    case "ellipse":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.ellipse, at),
        primitive.color,
      );
    case "rect":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, at),
        primitive.color,
      );
    case "triangle":
      return addTriangle(shapes, primitive, at);
    case "blockArc":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.blockArc, at),
        primitive.color,
      );
    case "donut":
      return filled(
        shapes.addGeometricShape(PowerPoint.GeometricShapeType.donut, at),
        primitive.color,
      );
  }
}

// The Triangle preset stands on its box's bottom edge with the tip up; the part
// of an area under its baseline hangs from the top edge, which is the same
// shape turned half a turn (PowerPoint has no flip): turnShape writes the turn,
// once the shape has its name.
function addTriangle(
  shapes: PowerPoint.ShapeCollection,
  primitive: Triangle,
  at: Box,
): PowerPoint.Shape {
  return filled(
    shapes.addGeometricShape(PowerPoint.GeometricShapeType.triangle, at),
    primitive.color,
  );
}

// The half turn of a hanging triangle. It is written after the shape's name,
// in the same batch, so a host that refuses the rotation has already named the
// shape and the cleanup sweep by name finds it.
export function turnShape(shape: PowerPoint.Shape, primitive: Primitive): void {
  if (primitive.kind === "triangle" && primitive.down) {
    shape.rotation = HALF_TURN;
  }
}

// The tip position the host is given for a triangle: a half turn mirrors the
// tip, so a triangle drawn tip-down is written at the other corner.
export function triangleTip(triangle: Triangle): number {
  return triangle.down ? 1 - triangle.apex : triangle.apex;
}

// The adjustment points a primitive's preset is given, in the host's units
// (S1, 03.10.2026; never an OOXML number: the host stores a value x 100000 in
// an int32, so a raw 50000 overflows and draws at 0). An angle is degrees
// clockwise from 3 o'clock; every other point is a fraction.
export function adjustmentsOf(primitive: Primitive): number[] {
  switch (primitive.kind) {
    case "wedge":
      // A pie: the angle it starts at, the angle it ends at.
      return [primitive.start, primitive.end];
    case "triangle":
      // Where the tip sits along the top edge, of the box's width.
      return [triangleTip(primitive)];
    case "blockArc":
      // A doughnut slice: start angle, end angle, and the ring's thickness
      // as a fraction of the box's side.
      return [primitive.start, primitive.end, primitive.thickness];
    case "donut":
      // A whole ring: its thickness, of the box's side.
      return [primitive.thickness];
    default:
      return [];
  }
}

// Where a primitive's adjustments are written depends on the host, and costs no
// extra round trip on either (S1 2.1, the web rig of 05.10).
// - PowerPoint for the web takes them in the batch that adds the shape and must
//   be given them there: a proxy a LATER batch uses is re-anchored by id
//   (shapes.getItem), which the web cannot resolve on a slide the editor has
//   never shown ("InvalidParam passed to GetItem(id)", 5010 - a pie, doughnut
//   or area inserted there died right after its first chunk).
// - The desktop hands out a shape's adjustments only once it has heard of the
//   shape (Mac and Windows were not measured for the same batch; the later
//   order is the one every host took for a pie), so shapeAdjustments queues
//   them after the sync that added the shapes and they travel with the next
//   one - the following chunk's, or the grouping sync.
export function adjustsInAddBatch(host: DrawHost): boolean {
  return host === "web";
}

// One primitive's adjustment points, written on its shape's proxy. On the web
// the caller writes them after the shape's name (and its half turn), in the
// batch that adds the shape, so a host that refuses one has already named the
// shape and the cleanup sweep by name finds it.
export function adjustShape(
  shape: PowerPoint.Shape,
  primitive: Primitive,
): void {
  adjustmentsOf(primitive).forEach((value, index) => {
    shape.adjustments.set(index, value);
  });
}

// The desktop's order: every adjustment of a chunk the host has just confirmed.
export function shapeAdjustments(added: Added[]): void {
  for (const { shape, primitive } of added) adjustShape(shape, primitive);
}
