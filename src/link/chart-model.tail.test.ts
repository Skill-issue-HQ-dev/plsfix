// What a toast says after a link's label ("Inserted <label>", "Put <label> back"):
// the label and what the host had to say about it read as ONE sentence. A chart
// that falls back to a picture used to read "Inserted Combo chart. as a picture:
// <reason>" - a lower-case clause after a full stop (found on real Office for the
// web, 05.10.2026). A note that reads on from the label goes before the full
// stop; any other note is a sentence of its own after it. Pure: no Office.js.

import { describe, expect, it } from "vitest";
import {
  labelledNote,
  labelTail,
  needsNewerPowerPoint,
  NO_SHAPE_API,
  PICTURE_CHOSEN,
  pictureNote,
  SMOOTH_STRAIGHT,
  STACK_NEGATIVE,
} from "./chart-model";

const LABEL = "Inserted Combo chart";
const OVERLAP = "Placed over other objects: no free space on this slide";
const DECK = "The deck's copy of this link was not saved: the store is full";

// The whole line, the way main.ts and deck-actions.ts compose it.
const said = (...notes: string[]): string => `${LABEL}${labelTail(notes)}`;

describe("labelTail", () => {
  it("is the full stop alone for a link with nothing to say", () => {
    expect(labelTail([])).toBe(".");
    expect(said()).toBe("Inserted Combo chart.");
  });

  it("reads a picture's reason on from the label, as one sentence", () => {
    expect(said(pictureNote("this chart type is not drawn as shapes"))).toBe(
      "Inserted Combo chart as a picture: this chart type is not drawn as shapes.",
    );
    expect(said(NO_SHAPE_API)).toBe(
      "Inserted Combo chart as a picture: this PowerPoint has no shape API (Microsoft 365 2504 / Mac 16.96 or newer).",
    );
    expect(said(needsNewerPowerPoint("pie"))).toBe(
      "Inserted Combo chart as a picture: pie shapes need Microsoft 365 2601 / Mac 16.105 or newer.",
    );
  });

  it("puts a drawn chart's fragment in brackets, the way an Update all line does", () => {
    expect(said(SMOOTH_STRAIGHT)).toBe(
      "Inserted Combo chart (smooth lines are drawn straight).",
    );
    expect(said(STACK_NEGATIVE)).toBe(
      "Inserted Combo chart (values under zero in a stack are drawn approximately).",
    );
    expect(labelledNote("Combo chart", SMOOTH_STRAIGHT)).toBe(
      "Combo chart (smooth lines are drawn straight)",
    );
  });

  it("leaves a sentence of the host's own after the full stop, as it always did", () => {
    expect(said(OVERLAP)).toBe(`Inserted Combo chart. ${OVERLAP}`);
    expect(said(PICTURE_CHOSEN)).toBe(
      "Inserted Combo chart. Picture (your choice)",
    );
    expect(said(OVERLAP, PICTURE_CHOSEN)).toBe(
      `Inserted Combo chart. ${OVERLAP}. Picture (your choice)`,
    );
  });

  it("puts the clause before the full stop and the sentences after it, whatever order they came in", () => {
    const reason = pictureNote("7 series; shapes draw up to 6");
    expect(said(OVERLAP, reason)).toBe(
      `Inserted Combo chart as a picture: 7 series; shapes draw up to 6. ${OVERLAP}`,
    );
    expect(said(reason, OVERLAP)).toBe(said(OVERLAP, reason));
    expect(said(OVERLAP, SMOOTH_STRAIGHT)).toBe(
      `Inserted Combo chart (smooth lines are drawn straight). ${OVERLAP}`,
    );
  });

  it("keeps a deck note that was joined to the clause, and doubles no full stop", () => {
    const reason = pictureNote("this chart type is not drawn as shapes");
    expect(said(`${reason}. ${DECK}`)).toBe(
      `Inserted Combo chart as a picture: this chart type is not drawn as shapes. ${DECK}.`,
    );
    expect(said(`${SMOOTH_STRAIGHT}. ${DECK}`)).toBe(
      `Inserted Combo chart (smooth lines are drawn straight). ${DECK}.`,
    );
    expect(said(pictureNote("one more sentence."))).toBe(
      "Inserted Combo chart as a picture: one more sentence.",
    );
  });

  it("never leaves a lower-case clause after a full stop, whatever the note", () => {
    const notes = [
      pictureNote("this chart type is not drawn as shapes"),
      NO_SHAPE_API,
      needsNewerPowerPoint("doughnut"),
      SMOOTH_STRAIGHT,
      STACK_NEGATIVE,
      `${SMOOTH_STRAIGHT}. ${DECK}`,
      OVERLAP,
      PICTURE_CHOSEN,
      DECK,
    ];
    for (const note of notes) {
      expect(said(note), note).not.toMatch(/\. [a-z]/);
      expect(said(OVERLAP, note), note).not.toMatch(/\. [a-z]/);
    }
  });
});
