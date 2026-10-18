// The room a value scale keeps over its tallest value for the label that stands
// on it: the same range with its top raised, so the highest value lands `room`
// points under the top of the plot instead of on it and its label stays out of
// the title band above the plot. Shared by the column, waterfall and combo
// layouts. Pure: no Office.js, and nothing imported at run time.
// Invariant: a plot gives up at most a third of its height, so a chart in a
// small box keeps a plot.

import type { ValueScale } from "./chart-shapes";

const ROOM_SHARE = 3;

// The scale with `headroom` points of the plot's `height` kept over its highest
// value. A range of nothing (every value zero) counts as the one the value maths
// reads it as, so even its zero line stands off the title.
export function withRoom(
  scale: ValueScale,
  height: number,
  headroom: number,
): ValueScale {
  const room = Math.min(headroom, height / ROOM_SHARE);
  if (height <= room) return scale;
  const range = scale.max - scale.min || 1;
  return {
    min: scale.min,
    max: scale.min + (range * height) / (height - room),
  };
}
