// End labels apart: the labels a series without data labels keeps over its last
// point all end in the same column of the plot, so those of lines that finish
// close together would sit on one another. spreadEnds pushes them apart, top to
// bottom, by just enough, keeps their order and keeps them inside the room they
// were given; labels over different categories never touch each other.

import { describe, expect, it } from "vitest";
import { LABEL_HEIGHT, type Primitive, type Text } from "./chart-shapes";
import { spreadEnds } from "./chart-shapes-ends";
import type { Box } from "./layout";

const BOUNDS: Box = { left: 0, top: 100, width: 400, height: 200 };
const H = LABEL_HEIGHT;

function text(name: string, top: number, left = 300): Text {
  return {
    kind: "text",
    box: { left, top, width: 30, height: H },
    text: name,
    size: 9,
    bold: false,
    color: "#111111",
    align: "c",
    name,
  };
}

const tops = (out: Primitive[], names: string[]): number[] =>
  names.map((name) => out.find((one) => one.name === name)!.box.top);

function spread(labels: Text[], columns: number[]): Primitive[] {
  const ends = labels.map((one, i) => ({
    name: one.name,
    column: columns[i]!,
  }));
  return spreadEnds(labels, ends, BOUNDS);
}

describe("spreadEnds", () => {
  it("leaves labels that do not touch exactly where they are", () => {
    const labels = [text("a", 120), text("b", 120 + H), text("c", 250)];
    expect(spread(labels, [5, 5, 5])).toEqual(labels);
  });

  it("pushes the lower of two touching labels down by just enough", () => {
    const out = spread([text("a", 150), text("b", 155)], [5, 5]);
    expect(tops(out, ["a", "b"])).toEqual([150, 150 + H]);
  });

  it("keeps the order: the label that was higher stays higher", () => {
    const out = spread([text("a", 155), text("b", 150)], [5, 5]);
    expect(tops(out, ["a", "b"])).toEqual([150 + H, 150]);
  });

  it("chains: three labels on one spot end up a label's height apart", () => {
    const out = spread(
      [text("a", 200), text("b", 200), text("c", 201)],
      [5, 5, 5],
    );
    expect(tops(out, ["a", "b", "c"])).toEqual([200, 200 + H, 200 + 2 * H]);
  });

  it("pushes a stack that would run off the bottom back up, order kept", () => {
    const floor = BOUNDS.top + BOUNDS.height - H;
    const out = spread([text("a", floor), text("b", floor - 2)], [5, 5]);
    expect(tops(out, ["b", "a"])).toEqual([floor - H, floor]);
  });

  it("keeps the first labels that fit and drops the rest, none overlapping", () => {
    const count = 14;
    const labels = Array.from({ length: count }, (_one, i) =>
      text(`l${String(i)}`, 150 + i),
    );
    const out = spread(
      labels,
      labels.map(() => 5),
    );
    const fit = Math.floor(BOUNDS.height / H);
    expect(fit).toBeLessThan(count);
    expect(out.map((one) => one.name)).toEqual(
      labels.slice(0, fit).map((one) => one.name),
    );
    for (const one of out) {
      expect(one.box.top).toBeGreaterThanOrEqual(BOUNDS.top);
      expect(one.box.top + one.box.height).toBeLessThanOrEqual(
        BOUNDS.top + BOUNDS.height,
      );
    }
    const sorted = tops(
      out,
      out.map((one) => one.name),
    );
    expect([...sorted].sort((a, b) => a - b)).toEqual(sorted);
    sorted.slice(1).forEach((top, k) => {
      expect(top - sorted[k]!).toBeGreaterThanOrEqual(H);
    });
  });

  it("drops the same labels every time: the order given decides, not the tops", () => {
    const labels = Array.from({ length: 14 }, (_one, i) =>
      text(`l${String(i)}`, 290 - i),
    );
    const ends = labels.map((one) => ({ name: one.name, column: 5 }));
    const first = spreadEnds(labels, ends, BOUNDS).map((one) => one.name);
    const again = spreadEnds([...labels], ends, BOUNDS).map((one) => one.name);
    expect(first).toEqual(labels.slice(0, 11).map((one) => one.name));
    expect(again).toEqual(first);
  });

  it("keeps at least one label of a column in a room too small for any", () => {
    const small: Box = { left: 0, top: 100, width: 400, height: H - 4 };
    const labels = [text("a", 100), text("b", 100)];
    const out = spreadEnds(
      labels,
      labels.map((one) => ({ name: one.name, column: 5 })),
      small,
    );
    expect(out.map((one) => one.name)).toEqual(["a"]);
  });

  it("works column by column: labels over other categories move nothing", () => {
    const out = spread(
      [text("a", 150), text("b", 152), text("c", 150, 100)],
      [5, 5, 3],
    );
    expect(tops(out, ["a", "b", "c"])).toEqual([150, 150 + H, 150]);
  });

  it("changes nothing but the top of a label it moves", () => {
    const labels = [text("a", 150), text("b", 152)];
    const out = spread(labels, [5, 5]);
    expect(out[1]).toEqual({
      ...labels[1],
      box: { ...labels[1]!.box, top: 150 + H },
    });
  });

  it("ignores a name that is not in the layout, and a column of one", () => {
    const labels = [text("a", 150), text("b", 152)];
    const out = spreadEnds(
      labels,
      [
        { name: "a", column: 1 },
        { name: "ghost", column: 1 },
        { name: "b", column: 2 },
      ],
      BOUNDS,
    );
    expect(out).toEqual(labels);
  });

  it("hands back primitives that are not labels as they were", () => {
    const rect: Primitive = {
      kind: "rect",
      box: { left: 0, top: 150, width: 10, height: 10 },
      color: "#000000",
      name: "bar",
    };
    const out = spreadEnds(
      [rect, text("a", 150), text("b", 152)],
      [
        { name: "a", column: 1 },
        { name: "b", column: 1 },
      ],
      BOUNDS,
    );
    expect(out[0]).toBe(rect);
  });
});
