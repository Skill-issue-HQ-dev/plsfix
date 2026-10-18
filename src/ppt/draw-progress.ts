// What a long chart draw tells the user and asks him: one line per finished
// chunk ("Drawing <label>: 48 of 92 shapes. Keep this tab in front.", the tab
// sentence on the web only, where a tab in the background never answers) and
// one question above an estimated two minutes: "Draw it" or "Insert as
// picture" - or, on a slide so full that not even one add can finish, "Insert
// as picture" or "Cancel" alone. The pane installs the toast as both; nothing
// here touches the DOM or Office.js beyond the host's name.
// Invariant: with no UI installed a draw is silent and never asks, and the
// only picture is the user's own answer.

import type { AskingToast, Choice } from "../ui/toast";
import {
  askMinutes,
  ASK_ABOVE_SECONDS,
  canDraw,
  estimateSeconds,
  SYNC_TIMEOUT_MS,
  type DrawHost,
} from "./chart-cost";
import { drawHost, type DrawProgress } from "./chart-draw";

// The sentence the web needs and the desktop does not: a PowerPoint for the
// web tab in the background never answers a write batch (lessons, 30.08).
const KEEP_TAB = "Keep this tab in front.";

// How long a progress line stays up. A chunk may take up to its host's sync
// deadline, and the line must still be there when the next one lands (half as
// long again as the longest deadline); the guard's own final toast replaces it
// the moment the action ends.
export const PROGRESS_STAY_MS =
  1.5 * Math.max(...Object.values(SYNC_TIMEOUT_MS));

export function progressMessage(
  label: string,
  done: number,
  total: number,
  host: DrawHost,
): string {
  const line = `Drawing ${label}: ${String(done)} of ${String(total)} shapes.`;
  return host === "web" ? `${line} ${KEEP_TAB}` : line;
}

// What the user decided about a long draw. "cancel" is also what a question
// that something else dismissed means: nothing is inserted.
export type DrawChoice = "draw" | "picture" | "cancel";

// What the pane gives a draw to talk through.
export interface DrawUi {
  // One line, up until the next one or the action's final toast replaces it.
  progress(message: string): void;
  // The long-draw question. A UI without it never asks, so the draw starts.
  ask?(
    question: string,
    choices: readonly Choice<DrawChoice>[],
  ): Promise<DrawChoice>;
}

let ui: DrawUi | null = null;

export function installDrawUi(next: DrawUi | null): void {
  ui = next;
}

// The pane's own toast as the UI: every line a success message held long
// enough to outlast the slowest chunk, and the question on the same banner.
export function toastDrawUi(toast: AskingToast): DrawUi {
  return {
    progress: (message) => {
      toast.show(message, "success", undefined, PROGRESS_STAY_MS);
    },
    ask: (question, choices) => toast.ask(question, choices, "cancel"),
  };
}

// The per-chunk callback of a draw labelled `label`, or none when no pane has
// installed a UI: chart-draw.ts then skips the call altogether.
export function drawProgress(label: string): DrawProgress | undefined {
  const sink = ui;
  if (sink === null) return undefined;
  return (done, total) => {
    sink.progress(progressMessage(label, done, total, drawHost()));
  };
}

// What an insert knows before it draws: the link's label, how many shapes the
// chart is, how many the target slide already holds (an add costs more for
// each) and whether the payload carries picture bytes to fall back on - a
// re-insert from the deck's own copy carries none.
export interface DrawRequest {
  label: string;
  shapes: number;
  existing: number;
  hasPicture: boolean;
}

// The question and its two answers for a draw estimated at `minutes`. A
// re-insert with no picture bytes can only be drawn or cancelled. Split out so
// the UX gate shows the very words the pane asks.
export function drawQuestion(
  request: DrawRequest,
  minutes: number,
): { question: string; choices: readonly Choice<DrawChoice>[] } {
  const min = String(minutes);
  const other: Choice<DrawChoice> = request.hasPicture
    ? { label: "Insert as picture", value: "picture" }
    : { label: "Cancel", value: "cancel" };
  const draw: Choice<DrawChoice> = {
    label: request.hasPicture ? `Draw it (about ${min} min)` : "Draw it",
    value: "draw",
  };
  return {
    question: `${request.label}: ${String(request.shapes)} shapes will take about ${min} min to draw here.`,
    choices: [draw, other],
  };
}

// The question for a slide that cannot take a chart at all: even one add is
// predicted past the safe line, so there is nothing to draw, only the picture
// (when the payload carries one) or nothing. Split out like drawQuestion.
export function tooFullQuestion(request: DrawRequest): {
  question: string;
  choices: readonly Choice<DrawChoice>[];
} {
  const cancel: Choice<DrawChoice> = { label: "Cancel", value: "cancel" };
  return {
    question: `${request.label}: this slide holds ${String(request.existing)} shapes, too many to draw a chart on here.`,
    choices: request.hasPicture
      ? [{ label: "Insert as picture", value: "picture" }, cancel]
      : [cancel],
  };
}

// Draw it, unless it is estimated to take over two minutes: then the user
// decides, and the picture is only ever his answer. On a slide too full for
// even one add the draw is not offered at all.
export async function chooseDrawing(request: DrawRequest): Promise<DrawChoice> {
  const sink = ui;
  if (sink?.ask === undefined) return "draw";
  const host = drawHost();
  if (!canDraw(request.existing, host)) {
    const full = tooFullQuestion(request);
    return sink.ask(full.question, full.choices);
  }
  const seconds = estimateSeconds(request.shapes, request.existing, host);
  if (seconds <= ASK_ABOVE_SECONDS) return "draw";
  const { question, choices } = drawQuestion(request, askMinutes(seconds));
  return sink.ask(question, choices);
}
