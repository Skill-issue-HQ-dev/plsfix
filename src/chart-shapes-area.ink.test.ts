// The ink of an area's labels: chosen once the end labels have been spread
// apart, against the fill that is really under the middle of each label (the
// last piece painted there, none = the ink), not against the fill of the
// series the label belongs to.

import { describe, expect, it } from "vitest";
import { darkFill, type Primitive, type Text } from "./chart-shapes";
import {
  BOX,
  DARK,
  INK,
  LIGHT,
  WHITE,
  areaOf,
  fills,
  named,
} from "../test/chart-area.support";
import { layoutArea } from "./chart-shapes-area";
import { fillUnder } from "./chart-shapes-area-hit";

const labels = (out: Primitive[]): Text[] => named(out, "label ") as Text[];

// What each label should be written in, judged where the label ended up.
function wanted(out: Primitive[], one: Text): string {
  const { left, top, width, height } = one.box;
  const under = fillUnder(fills(out), left + width / 2, top + height / 2);
  return under !== null && darkFill(under) ? WHITE : INK;
}

describe("layoutArea: the ink of a label", () => {
  it("is white for an end label that lands on a dark fill of another series", () => {
    // Both series end on the last column; the orange series' 7 sits above its
    // own line, which is over the navy series' fill.
    const data = areaOf("area", [
      { values: [1, 2, 10], color: DARK, extra: { dataLabels: false } },
      { values: [1, 2, 7], color: LIGHT, extra: { dataLabels: false } },
    ]);
    const out = layoutArea(data, BOX);
    const seven = labels(out).find((one) => one.text === "7")!;
    expect(seven.color).toBe(WHITE);
    for (const one of labels(out)) expect(one.color).toBe(wanted(out, one));
  });

  it.each([
    {
      values: [
        [1, 9, 5],
        [2, 8, 5],
      ],
    },
    {
      values: [
        [1, 2, 10],
        [1, 2, 7],
      ],
    },
    {
      values: [
        [3, 3, 3],
        [3, 3, 3],
        [3, 3, 3],
      ],
    },
  ])(
    "each plain-area end label contrasts with the fill actually under it: %o",
    ({ values }) => {
      const data = areaOf(
        "area",
        values.map((row, i) => ({
          values: row,
          color: i % 2 === 0 ? DARK : LIGHT,
          extra: { dataLabels: false },
        })),
      );
      const out = layoutArea(data, BOX);
      expect(labels(out).length).toBe(values.length);
      for (const one of labels(out)) expect(one.color).toBe(wanted(out, one));
    },
  );

  it("is the plain ink where no fill lies under the label", () => {
    const data = areaOf("area", [
      { values: [10, 40, 20], color: DARK, extra: { dataLabels: false } },
    ]);
    const [only] = labels(layoutArea(data, BOX));
    // Above the line, over the empty plot.
    expect(only!.color).toBe(INK);
  });

  it("takes a value label inside a light fill over a dark one for the ink", () => {
    const data = areaOf("area", [
      { values: [30, 30, 30], color: DARK },
      { values: [10, 20, 10], color: LIGHT },
    ]);
    const out = layoutArea(data, BOX);
    for (const one of labels(out)) expect(one.color).toBe(wanted(out, one));
    // The light series' 20 rides above its line, inside the dark series' fill.
    const twenty = labels(out).find((one) => one.text === "20")!;
    expect(twenty.color).toBe(WHITE);
  });
});
