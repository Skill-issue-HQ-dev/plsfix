// Can THIS host draw a deck's copy of a link again? Put back, its row's button
// and the Inbox's hide set all take the answer from here. A "no" says why in the
// kind's own sentence, never "this PowerPoint cannot" for a chart no host draws.

import { afterEach, describe, expect, it, vi } from "vitest";
import type { ChartKind } from "../link/chart-model";
import type { ChartSeries } from "../link/chart-series";
import { identityOnly, type DeckRecord } from "../link/deck-record";
import {
  CHART,
  LINK,
  picture,
  record,
  table,
  TEXT,
} from "../link/deck-record.support";
import { drawingFrom, recordedLinkIds } from "./deck-drawable";

// A host that answers only for the API sets it names: Office.js's own check.
function hostWith(...sets: string[]): void {
  vi.stubGlobal("Office", {
    context: {
      requirements: {
        isSetSupported: (name: string, version: string) =>
          name === "PowerPointApi" && sets.includes(version),
      },
    },
  });
}

afterEach(() => {
  vi.unstubAllGlobals();
});

const chartOf = (kind: ChartKind): DeckRecord =>
  record(picture({ chart: { ...CHART, kind } }), { ...LINK, kind: "chart" });

describe("a record the host can draw again", () => {
  it("is a text, a table, or a chart the host draws as shapes", () => {
    hostWith("1.8", "1.10");
    for (const made of [record(TEXT), record(table(2, 2)), chartOf("column")]) {
      const drawing = drawingFrom(made);
      expect(drawing.ok).toBe(true);
    }
  });

  it("hands over the payload the record stands for, a chart's picture bytes empty", () => {
    hostWith("1.8", "1.10");
    const drawing = drawingFrom(chartOf("column"));
    expect(drawing).toMatchObject({
      ok: true,
      payload: { kind: "picture", png: "", chart: { kind: "column" } },
    });
  });

  it("needs no host check for a text or a table: they are drawn on every host", () => {
    hostWith();
    expect(drawingFrom(record(TEXT)).ok).toBe(true);
    expect(drawingFrom(record(table(2, 2))).ok).toBe(true);
  });
});

describe("a record the host cannot draw again", () => {
  it("is one that kept the identity only: no data, and it says to send it again", () => {
    hostWith("1.8", "1.10");
    const drawing = drawingFrom(identityOnly(record(TEXT)));
    expect(drawing).toMatchObject({ ok: false, kept: false });
    expect(drawing).toMatchObject({
      issue: "this deck kept no data for it",
      way: "Send it again from Excel.",
    });
  });

  it("is a chart kind no shape version exists for: the pending-kind sentence, the data kept", () => {
    hostWith("1.8", "1.10");
    const drawing = drawingFrom(chartOf("scatter"));
    expect(drawing).toMatchObject({
      ok: false,
      kept: true,
      issue: "scatter charts are not drawn as shapes yet",
    });
    if (drawing.ok) throw new Error("a scatter is not drawn as shapes");
    // The way back is the Inbox, which inserts it as a picture.
    expect(drawing.way).toBe(
      "Send it again from Excel: the Inbox inserts it as a picture.",
    );
    expect(`${drawing.issue} ${drawing.way}`).not.toMatch(/this PowerPoint/);
  });

  it("is a combo with a series the slide has no shapes for: the whole chart is judged, not its kind", () => {
    hostWith("1.8", "1.10");
    const base = CHART.series[0]!;
    const combo = (second: ChartSeries): DeckRecord =>
      record(
        picture({
          chart: {
            ...CHART,
            kind: "combo",
            series: [{ ...base, type: "column" }, second],
          },
        }),
        { ...LINK, kind: "chart" },
      );
    const margin = { ...base, name: "Margin", type: "line" as const };
    expect(drawingFrom(combo({ ...margin, axis: "secondary" })).ok).toBe(true);
    expect(
      drawingFrom(combo({ ...base, name: "Units", type: "bar" })),
    ).toMatchObject({
      ok: false,
      kept: true,
      issue: "bar series in a combo chart are not drawn as shapes",
    });
  });

  it("is a waterfall too short to bridge: its point count in the issue, the data kept", () => {
    hostWith("1.8", "1.10");
    // CHART has two categories, a bridge needs three.
    expect(drawingFrom(chartOf("waterfall"))).toMatchObject({
      ok: false,
      kept: true,
      issue: "2 points; waterfall shapes need at least 3",
    });
  });

  it("is a chart whose kind needs a newer PowerPoint than this one", () => {
    hostWith("1.8");
    expect(drawingFrom(chartOf("column")).ok).toBe(true);
    expect(drawingFrom(chartOf("pie"))).toMatchObject({
      ok: false,
      kept: true,
      issue: "pie shapes need Microsoft 365 2601 / Mac 16.105 or newer",
    });
  });

  it("is any chart on a PowerPoint with no shape API, in that case's own words", () => {
    hostWith();
    expect(drawingFrom(chartOf("column"))).toMatchObject({
      ok: false,
      kept: true,
      issue:
        "this PowerPoint has no shape API (Microsoft 365 2504 / Mac 16.96 or newer)",
    });
  });
});

describe("the links the Inbox leaves to the list", () => {
  const held = (...made: DeckRecord[]): Map<string, DeckRecord> =>
    new Map(made.map((one) => [one.id, one]));
  const idOf = (n: number): string => n.toString(16).padStart(32, "0");
  const as = (n: number, made: DeckRecord): DeckRecord => ({
    ...made,
    id: idOf(n),
  });

  it("are the ones the list can put back: a record that cannot be drawn keeps its Inbox row", () => {
    hostWith("1.8", "1.10");
    const ids = recordedLinkIds(
      held(
        as(1, record(TEXT)),
        as(2, identityOnly(record(TEXT))),
        as(3, chartOf("scatter")),
        as(4, chartOf("column")),
      ),
    );
    expect([...ids]).toEqual([idOf(1), idOf(4)]);
  });

  it("follow the host: the same chart is hidden where it can be drawn and shown where it cannot", () => {
    const records = held(as(1, chartOf("pie")));
    hostWith("1.8", "1.10");
    expect(recordedLinkIds(records).size).toBe(1);
    hostWith("1.8");
    expect(recordedLinkIds(records).size).toBe(0);
  });

  it("are none for a deck with no records", () => {
    expect(recordedLinkIds(new Map()).size).toBe(0);
  });
});
