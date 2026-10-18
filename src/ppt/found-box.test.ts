// The box a scanned link is found at: the host's own, or - for a group PowerPoint
// for Mac reports with no area - the union of the shapes the scan found inside
// it, down through sub-groups, and only those: another group's shapes, or
// another slide's, are not its own.

import { describe, expect, it } from "vitest";
import type { Box } from "../layout";
import { scannedBox } from "./found-box";
import type { PlacedShape } from "./shapes";

function placed(
  id: string,
  type: string,
  box: Box,
  groupPath: string[] = [],
  slideId = "slide-1",
): PlacedShape {
  return {
    slideId,
    slideIndex: 0,
    shape: { id, type, ...box } as unknown as PowerPoint.Shape,
    groupPath,
  };
}

const NONE: Box = { left: 0, top: 0, width: 0, height: 0 };
const LEFT: Box = { left: 100, top: 100, width: 200, height: 120 };
const RIGHT: Box = { left: 320, top: 90, width: 200, height: 150 };

describe("scannedBox", () => {
  it("is the host's own box for a shape that has one", () => {
    const picture = placed("p", "GeometricShape", LEFT);
    expect(scannedBox(picture, [picture])).toEqual(LEFT);
  });

  it("is the group's own box when the host reports one", () => {
    const group = placed("g", "Group", { ...LEFT, width: 500 });
    const inside = placed("a", "GeometricShape", RIGHT, ["g"]);
    expect(scannedBox(group, [group, inside]).width).toBe(500);
  });

  it("is the union of what is inside a group the host reports with no box", () => {
    const group = placed("g", "Group", NONE);
    const all = [
      group,
      placed("a", "GeometricShape", LEFT, ["g"]),
      placed("b", "GeometricShape", RIGHT, ["g"]),
    ];
    expect(scannedBox(group, all)).toEqual({
      left: 100,
      top: 90,
      width: 420,
      height: 150,
    });
  });

  it("reads through sub-groups that report no box either", () => {
    const group = placed("g", "Group", NONE);
    const tier = placed("t", "Group", NONE, ["g"]);
    const all = [
      group,
      tier,
      placed("a", "GeometricShape", LEFT, ["g", "t"]),
      placed("b", "GeometricShape", RIGHT, ["g"]),
    ];
    expect(scannedBox(group, all)).toEqual({
      left: 100,
      top: 90,
      width: 420,
      height: 150,
    });
  });

  it("counts only its own children: not another group's, not another slide's", () => {
    const group = placed("g", "Group", NONE);
    const all = [
      group,
      placed("other", "Group", LEFT),
      placed("a", "GeometricShape", LEFT, ["g"]),
      placed("b", "GeometricShape", RIGHT, ["other"]),
      placed("c", "GeometricShape", RIGHT, ["g"], "slide-2"),
    ];
    expect(scannedBox(group, all)).toEqual(LEFT);
  });

  it("stays the host's box when there is nothing inside to read", () => {
    const group = placed("g", "Group", NONE);
    expect(scannedBox(group, [group])).toEqual(NONE);
    const rule = placed("r", "Line", { ...NONE, width: 300 });
    expect(scannedBox(rule, [rule])).toEqual({ ...NONE, width: 300 });
  });
});
