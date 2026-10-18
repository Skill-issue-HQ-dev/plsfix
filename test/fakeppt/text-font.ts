// PowerPoint.ShapeFont over the shape's own font: name, size, color, bold and
// italic. The fake keeps one run per shape, so the range's font is the shape's.
// Every write goes through writeFormat: a shape that refuses halts the batch, any
// other write is journaled. Invariant: a font field nobody set reads undefined.

import { ShapeFormatBound } from "./format-bound";
import { writeFormat } from "./format-write";
import type { FakeShapeFont } from "./model";

export class TextFontProxy extends ShapeFormatBound {
  private set<K extends keyof FakeShapeFont>(
    key: K,
    value: FakeShapeFont[K],
  ): void {
    writeFormat(this.deck, this.shapeId, "text", this.shape().font, key, value);
  }

  get name(): string | undefined {
    return this.shape().font.name;
  }
  set name(value: string | undefined) {
    this.set("name", value);
  }
  get size(): number | undefined {
    return this.shape().font.size;
  }
  set size(value: number | undefined) {
    this.set("size", value);
  }
  get color(): string | undefined {
    return this.shape().font.color;
  }
  set color(value: string | undefined) {
    this.set("color", value);
  }
  get bold(): boolean | undefined {
    return this.shape().font.bold;
  }
  set bold(value: boolean | undefined) {
    this.set("bold", value);
  }
  get italic(): boolean | undefined {
    return this.shape().font.italic;
  }
  set italic(value: boolean | undefined) {
    this.set("italic", value);
  }
}
