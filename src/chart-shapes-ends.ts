// End labels apart. A series without data labels keeps the value of its last
// point, so every such series ends in the same column of the plot and the
// labels of lines that finish close together would sit on one another. This
// pushes them apart, top to bottom, by just enough, keeping their order and
// the room they were given; labels over different categories never touch each
// other. A column with more labels than its room holds keeps the first ones, in
// the order given, and drops the rest: the legend carries identity. Pure: no
// Office.js. Invariant: only the top of a label moves, no two ever overlap.

import { LABEL_HEIGHT, type Primitive, type Text } from "./chart-shapes";
import type { Box } from "./layout";

// One end label: its name and the category (the column of the plot) it is over.
export interface End {
  name: string;
  column: number;
}

function isText(one: Primitive | undefined): one is Text {
  return one !== undefined && one.kind === "text";
}

// How many labels one column of the room holds, always at least one.
function capacity(room: Box): number {
  return Math.max(1, Math.floor(room.height / LABEL_HEIGHT));
}

// The tops of one column's labels, in the order given (no more than the room
// holds): stacked downwards a label's height apart where they touch, and the
// stack moved up where it would run off the bottom.
function separate(tops: number[], room: Box): number[] {
  const order = tops
    .map((_top, i) => i)
    .sort((a, b) => tops[a]! - tops[b]! || a - b);
  const placed = order.map((i) => tops[i]!);
  const ceiling = room.top + room.height - LABEL_HEIGHT;
  for (let k = 1; k < placed.length; k += 1) {
    placed[k] = Math.max(placed[k]!, placed[k - 1]! + LABEL_HEIGHT);
  }
  for (let k = placed.length - 1; k >= 0; k -= 1) {
    const below =
      k === placed.length - 1 ? ceiling : placed[k + 1]! - LABEL_HEIGHT;
    placed[k] = Math.min(placed[k]!, below);
  }
  const out = [...tops];
  order.forEach((i, k) => {
    out[i] = Math.max(placed[k]!, room.top);
  });
  return out;
}

// The layout with the end labels of each column pushed apart inside `room`,
// those past what a column holds left out. A name that is not in the layout, or
// a column of one label, changes nothing.
export function spreadEnds(
  primitives: Primitive[],
  ends: End[],
  room: Box,
): Primitive[] {
  const at = new Map(primitives.map((one, i) => [one.name, i] as const));
  const columns = new Map<number, number[]>();
  for (const { name, column } of ends) {
    const index = at.get(name);
    if (index === undefined || !isText(primitives[index])) continue;
    columns.set(column, [...(columns.get(column) ?? []), index]);
  }
  const out = [...primitives];
  const dropped = new Set<number>();
  for (const all of columns.values()) {
    if (all.length < 2) continue;
    const indices = all.slice(0, capacity(room));
    for (const index of all.slice(indices.length)) dropped.add(index);
    const labels = indices.map((i) => primitives[i] as Text);
    const tops = separate(
      labels.map((one) => one.box.top),
      room,
    );
    indices.forEach((index, k) => {
      const one = labels[k]!;
      out[index] = { ...one, box: { ...one.box, top: tops[k]! } };
    });
  }
  return out.filter((_one, i) => !dropped.has(i));
}
