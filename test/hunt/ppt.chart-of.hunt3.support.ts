// Shared fixtures for the H4 hunt tests (round 1): a ChartData of any drawable
// kind built from a few numbers, so a flow test can seed the relay with a
// chart of that kind without restating the payload. Owns no assertions.

import type { ChartData, ChartKind } from "../../src/link/chart-model";

export const PALETTE = ["#B27E54", "#14213D", "#2EC4B6", "#E63946", "#8AB17D"];

// One chart of `kind` over `points` categories and `seriesCount` series. A pie
// and a waterfall are one series; a combo alternates columns and lines; a
// scatter carries an x. Values are plain positives so every kind has bars,
// slices or segments to draw.
export function chartOf(
  kind: ChartKind,
  points: number,
  seriesCount = 1,
  extra: Partial<ChartData> = {},
): ChartData {
  const categories = Array.from({ length: points }, (_, i) => `Cat ${i + 1}`);
  const series = Array.from({ length: seriesCount }, (_, j) => {
    const values = categories.map((_name, i) => 20 + 7 * ((i + j) % 5));
    return {
      name: `Series ${j + 1}`,
      values,
      labels: values.map(String),
      colors: categories.map((_name, i) => PALETTE[(i + j) % PALETTE.length]!),
      ...(kind === "combo"
        ? {
            type: j % 2 === 0 ? ("column" as const) : ("line" as const),
          }
        : {}),
      ...(kind === "scatter" ? { x: values.map((_v, i) => i + 1) } : {}),
    };
  });
  return {
    v: 1,
    kind,
    title: "Revenue",
    categories,
    series,
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
    ...extra,
  };
}
