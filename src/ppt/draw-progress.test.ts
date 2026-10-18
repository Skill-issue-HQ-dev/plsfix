// What a long chart draw tells the user: the progress line per chunk (with the
// tab sentence on the web only) and the one sink the pane installs to show it.

import { afterEach, describe, expect, it, vi } from "vitest";
import {
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
} from "../../test/fakeppt";
import type { AskingToast } from "../ui/toast";
import {
  askMinutes,
  canDraw,
  estimateSeconds,
  SYNC_TIMEOUT_MS,
} from "./chart-cost";
import {
  chooseDrawing,
  drawProgress,
  drawQuestion,
  installDrawUi,
  PROGRESS_STAY_MS,
  progressMessage,
  toastDrawUi,
  tooFullQuestion,
  type DrawChoice,
  type DrawUi,
} from "./draw-progress";

afterEach(() => {
  installDrawUi(null);
  uninstallFakePpt();
});

function hostOn(platform: string): FakePptHelpers {
  const { helpers } = installFakePpt({ slides: 1 });
  helpers.setPlatform(platform);
  return helpers;
}

describe("the progress line", () => {
  it("says how far the draw is and, on the web, to keep the tab in front", () => {
    expect(progressMessage("Model: B4:F12", 48, 92, "web")).toBe(
      "Drawing Model: B4:F12: 48 of 92 shapes. Keep this tab in front.",
    );
  });

  it("leaves the tab sentence out on the desktop", () => {
    expect(progressMessage("Model: B4:F12", 48, 92, "desktop")).toBe(
      "Drawing Model: B4:F12: 48 of 92 shapes.",
    );
  });

  it("stays up longer than the slowest chunk any host's deadline allows", () => {
    // A chunk may take up to its host's sync deadline; the line must still be
    // there when the next one lands. Derived from the deadlines, not typed.
    const slowest = Math.max(...Object.values(SYNC_TIMEOUT_MS));
    expect(PROGRESS_STAY_MS).toBeGreaterThan(slowest);
    expect(PROGRESS_STAY_MS).toBe(1.5 * slowest);
  });
});

describe("the pane's toast as the sink", () => {
  it("shows each line as a success message that stays for the long hold", () => {
    const show = vi.fn();
    toastDrawUi({ show, ask: vi.fn() }).progress("Drawing X: 12 of 92 shapes.");
    expect(show).toHaveBeenCalledWith(
      "Drawing X: 12 of 92 shapes.",
      "success",
      undefined,
      PROGRESS_STAY_MS,
    );
  });

  it("asks through the toast, with cancel as what a dismissed question means", async () => {
    const ask = vi.fn(() => Promise.resolve("picture"));
    const ui = toastDrawUi({
      show: vi.fn(),
      ask: ask as unknown as AskingToast["ask"],
    });
    const choices = [{ label: "Draw it", value: "draw" as const }];

    await expect(ui.ask!("Draw?", choices)).resolves.toBe("picture");
    expect(ask).toHaveBeenCalledWith("Draw?", choices, "cancel");
  });
});

// The question above two minutes. By S1's cost table the 92-shape column chart
// is about 110 s on an empty web slide, so it does not ask, and about 130 s on
// a slide holding ten shapes, so it does; 125 shapes on that slide are about
// 220 s.
describe("the long-draw question", () => {
  const LONG = {
    label: "Model: B4:F12",
    shapes: 125,
    existing: 10,
    hasPicture: true,
  };

  // A UI whose question the user answers `value` to, recording what it asked.
  function answering(value: DrawChoice) {
    const ask = vi.fn<NonNullable<DrawUi["ask"]>>(() => Promise.resolve(value));
    installDrawUi({ progress: () => undefined, ask });
    return ask;
  }

  it("draws without a word when no pane can ask", async () => {
    hostOn("OfficeOnline");
    installDrawUi({ progress: () => undefined });
    await expect(chooseDrawing(LONG)).resolves.toBe("draw");
  });

  it("does not ask about a draw expected to take under two minutes", async () => {
    hostOn("OfficeOnline");
    const ask = answering("picture");
    await expect(
      chooseDrawing({ ...LONG, shapes: 92, existing: 0 }),
    ).resolves.toBe("draw");
    expect(ask).not.toHaveBeenCalled();
  });

  it("does not ask on the desktop, where the same chart is quick", async () => {
    hostOn("PC");
    const ask = answering("picture");
    await expect(chooseDrawing(LONG)).resolves.toBe("draw");
    expect(ask).not.toHaveBeenCalled();
  });

  it("asks above two minutes: Draw it with the minutes, or Insert as picture", async () => {
    hostOn("OfficeOnline");
    const ask = answering("picture");
    const minutes = askMinutes(estimateSeconds(125, 10, "web"));

    await expect(chooseDrawing(LONG)).resolves.toBe("picture");

    expect(ask).toHaveBeenCalledTimes(1);
    expect(ask.mock.calls[0]).toEqual([
      `Model: B4:F12: 125 shapes will take about ${String(minutes)} min to draw here.`,
      [
        { label: `Draw it (about ${String(minutes)} min)`, value: "draw" },
        { label: "Insert as picture", value: "picture" },
      ],
    ]);
  });

  it("offers Draw it or Cancel when the payload carries no picture", async () => {
    hostOn("OfficeOnline");
    const ask = answering("draw");
    const minutes = askMinutes(estimateSeconds(125, 10, "web"));

    await expect(chooseDrawing({ ...LONG, hasPicture: false })).resolves.toBe(
      "draw",
    );

    expect(ask.mock.calls[0]![1]).toEqual([
      { label: "Draw it", value: "draw" },
      { label: "Cancel", value: "cancel" },
    ]);
    expect(ask.mock.calls[0]![0]).toContain(`about ${String(minutes)} min`);
  });

  it("words the question without a host, for the UX gate to show", () => {
    expect(drawQuestion(LONG, 3)).toEqual({
      question: "Model: B4:F12: 125 shapes will take about 3 min to draw here.",
      choices: [
        { label: "Draw it (about 3 min)", value: "draw" },
        { label: "Insert as picture", value: "picture" },
      ],
    });
    expect(drawQuestion({ ...LONG, hasPicture: false }, 3).choices).toEqual([
      { label: "Draw it", value: "draw" },
      { label: "Cancel", value: "cancel" },
    ]);
  });

  it("hands back whatever the user answered", async () => {
    hostOn("OfficeOnline");
    for (const answer of ["draw", "picture", "cancel"] as const) {
      answering(answer);
      await expect(chooseDrawing(LONG)).resolves.toBe(answer);
    }
  });

  it("counts the shapes already on the slide: ten of them tip the 92-shape chart over", async () => {
    hostOn("OfficeOnline");
    const ask = answering("draw");
    await chooseDrawing({ ...LONG, shapes: 92, existing: 0 });
    expect(ask).not.toHaveBeenCalled();
    await chooseDrawing({ ...LONG, shapes: 92, existing: 10 });
    expect(ask).toHaveBeenCalledTimes(1);
    expect(ask.mock.calls[0]![0]).toContain("92 shapes");
  });

  it("names the minutes S1's table gives for the slide as it stands", async () => {
    hostOn("OfficeOnline");
    const ask = answering("draw");
    await chooseDrawing({ ...LONG, shapes: 92, existing: 10 });
    const minutes = askMinutes(estimateSeconds(92, 10, "web"));
    expect(minutes).toBe(2);
    expect(ask.mock.calls[0]![0]).toBe(
      `Model: B4:F12: 92 shapes will take about ${String(minutes)} min to draw here.`,
    );
  });
});

// A slide so full that even one add is predicted past the safe line: no
// chunking helps, so "Draw it" is not on offer, only the picture or nothing.
describe("a slide too full to draw on", () => {
  const FULL = {
    label: "Model: B4:F12",
    shapes: 92,
    existing: 1_000,
    hasPicture: true,
  };

  it("is where canDraw says no, about 940 shapes on the web", () => {
    expect(canDraw(FULL.existing, "web")).toBe(false);
  });

  it("offers Insert as picture or Cancel, never Draw it", async () => {
    hostOn("OfficeOnline");
    const ask = vi.fn<NonNullable<DrawUi["ask"]>>(() =>
      Promise.resolve("picture"),
    );
    installDrawUi({ progress: () => undefined, ask });

    await expect(chooseDrawing(FULL)).resolves.toBe("picture");

    expect(ask).toHaveBeenCalledTimes(1);
    expect(ask.mock.calls[0]).toEqual([
      "Model: B4:F12: this slide holds 1000 shapes, too many to draw a chart on here.",
      [
        { label: "Insert as picture", value: "picture" },
        { label: "Cancel", value: "cancel" },
      ],
    ]);
  });

  it("offers Cancel alone when the payload carries no picture", async () => {
    hostOn("OfficeOnline");
    const ask = vi.fn<NonNullable<DrawUi["ask"]>>(() =>
      Promise.resolve("cancel"),
    );
    installDrawUi({ progress: () => undefined, ask });

    await expect(chooseDrawing({ ...FULL, hasPicture: false })).resolves.toBe(
      "cancel",
    );

    expect(ask.mock.calls[0]![1]).toEqual([
      { label: "Cancel", value: "cancel" },
    ]);
  });

  it("asks even for a chart that would be quick on an emptier slide", async () => {
    hostOn("OfficeOnline");
    const ask = vi.fn<NonNullable<DrawUi["ask"]>>(() =>
      Promise.resolve("cancel"),
    );
    installDrawUi({ progress: () => undefined, ask });

    await chooseDrawing({ ...FULL, shapes: 3 });

    expect(ask).toHaveBeenCalledTimes(1);
  });

  it("is no problem on the desktop, which charges almost nothing for a full slide", async () => {
    hostOn("PC");
    const ask = vi.fn<NonNullable<DrawUi["ask"]>>(() =>
      Promise.resolve("cancel"),
    );
    installDrawUi({ progress: () => undefined, ask });

    await expect(chooseDrawing(FULL)).resolves.toBe("draw");
    expect(ask).not.toHaveBeenCalled();
  });

  it("draws without a word when no pane can ask, the host's deadline being the last word", async () => {
    hostOn("OfficeOnline");
    installDrawUi({ progress: () => undefined });
    await expect(chooseDrawing(FULL)).resolves.toBe("draw");
  });

  it("words the question without a host, for the UX gate to show", () => {
    expect(tooFullQuestion(FULL)).toEqual({
      question:
        "Model: B4:F12: this slide holds 1000 shapes, too many to draw a chart on here.",
      choices: [
        { label: "Insert as picture", value: "picture" },
        { label: "Cancel", value: "cancel" },
      ],
    });
    expect(tooFullQuestion({ ...FULL, hasPicture: false }).choices).toEqual([
      { label: "Cancel", value: "cancel" },
    ]);
  });
});

describe("the sink", () => {
  it("has nothing to say until a pane installs one", () => {
    hostOn("OfficeOnline");
    expect(drawProgress("Model: B4:F12")).toBeUndefined();
  });

  it("hands every chunk to the installed sink, worded for this host", () => {
    const lines: string[] = [];
    installDrawUi({ progress: (message) => lines.push(message) });
    hostOn("OfficeOnline");
    const report = drawProgress("Model: B4:F12")!;
    report(0, 92);
    report(12, 92);
    expect(lines).toEqual([
      "Drawing Model: B4:F12: 0 of 92 shapes. Keep this tab in front.",
      "Drawing Model: B4:F12: 12 of 92 shapes. Keep this tab in front.",
    ]);
  });

  it("words the line for the desktop when the host is one", () => {
    const lines: string[] = [];
    installDrawUi({ progress: (message) => lines.push(message) });
    hostOn("PC");
    drawProgress("Model: B4:F12")!(24, 92);
    expect(lines).toEqual(["Drawing Model: B4:F12: 24 of 92 shapes."]);
  });

  it("is silent again once the pane takes its sink away", () => {
    installDrawUi({ progress: () => undefined });
    installDrawUi(null);
    hostOn("PC");
    expect(drawProgress("Model: B4:F12")).toBeUndefined();
  });
});
