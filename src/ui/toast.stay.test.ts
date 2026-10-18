// @vitest-environment jsdom
// How long the one banner keeps a message (the default rule, no test override):
// a success stays 4 s + 70 ms a character, at most 12 s; an error stays until
// the next pane action takes it off (dismissError, the guard's call) or a
// click on the toast; a question stays until answered, as before.

import { afterEach, describe, expect, it, vi } from "vitest";
import { createToast } from "./toast";

afterEach(() => {
  vi.useRealTimers();
  document.body.replaceChildren();
  window.getSelection()?.removeAllRanges();
});

function attached(): HTMLElement {
  const root = document.createElement("div");
  document.body.append(root);
  return root;
}

const shown = (root: HTMLElement): boolean =>
  root.classList.contains("visible");

describe("a success toast", () => {
  it("stays four seconds plus 70 ms for every character", () => {
    vi.useFakeTimers();
    const root = attached();
    createToast(root).show("Saved");
    vi.advanceTimersByTime(4_349);
    expect(shown(root)).toBe(true);
    vi.advanceTimersByTime(2);
    expect(shown(root)).toBe(false);
    expect(root.className).toBe("toast");
  });

  it("stays longer for a longer sentence, up to twelve seconds", () => {
    vi.useFakeTimers();
    const root = attached();
    createToast(root).show("x".repeat(60));
    vi.advanceTimersByTime(8_199);
    expect(shown(root)).toBe(true);
    vi.advanceTimersByTime(2);
    expect(shown(root)).toBe(false);

    createToast(root).show("y".repeat(300));
    vi.advanceTimersByTime(11_999);
    expect(shown(root)).toBe(true);
    vi.advanceTimersByTime(2);
    expect(shown(root)).toBe(false);
  });

  it("restarts its clock when the next message replaces it", () => {
    vi.useFakeTimers();
    const root = attached();
    const toast = createToast(root);
    toast.show("Saved");
    vi.advanceTimersByTime(4_000);
    toast.show("Saved");
    vi.advanceTimersByTime(4_000);
    expect(shown(root)).toBe(true);
    vi.advanceTimersByTime(400);
    expect(shown(root)).toBe(false);
  });

  it("keeps the time a caller asks for, not the rule's", () => {
    vi.useFakeTimers();
    const root = attached();
    createToast(root).show(
      "Drawing: 12 of 92 shapes.",
      "success",
      undefined,
      500,
    );
    vi.advanceTimersByTime(501);
    expect(shown(root)).toBe(false);
  });
});

describe("an error toast", () => {
  it("is still showing after ten minutes", () => {
    vi.useFakeTimers();
    const root = attached();
    createToast(root).show("Excel is not connected.", "error");
    vi.advanceTimersByTime(600_000);
    expect(shown(root)).toBe(true);
    expect(root.className).toContain("error");
    expect(root.textContent).toContain("Excel is not connected.");
  });

  it("goes when it is clicked", () => {
    const root = attached();
    createToast(root).show("Excel is not connected.", "error");
    root
      .querySelector(".toast-text")!
      .dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(shown(root)).toBe(false);
    expect(root.className).toBe("toast");
  });

  it("is replaced by the next message, never stacked", () => {
    const root = attached();
    const toast = createToast(root);
    toast.show("First failure.", "error");
    toast.show("Done.");
    expect(root.querySelectorAll(".toast-text")).toHaveLength(1);
    expect(root.textContent).toContain("Done.");
    expect(root.className).toContain("success");
  });

  it("does not go when Copy details is pressed: that button copies", async () => {
    const root = attached();
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root).show("Failed", "error", "stack...");
    root.querySelector("button")!.click();
    expect(write).toHaveBeenCalledWith("stack...");
    await Promise.resolve();
    await Promise.resolve();
    expect(shown(root)).toBe(true);
    expect(root.textContent).toContain("Details copied.");
  });

  it("stays when the click ended a selection of its own words", () => {
    const root = attached();
    createToast(root).show("Copy failed: select the text.", "error");
    const text = root.querySelector(".toast-text")!;
    const range = document.createRange();
    range.selectNodeContents(text);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);

    text.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(shown(root)).toBe(true);

    selection.removeAllRanges();
    text.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(shown(root)).toBe(false);
  });

  it("an error given its own time still leaves on schedule", () => {
    vi.useFakeTimers();
    const root = attached();
    createToast(root).show("Heads up.", "error", undefined, 1_000);
    vi.advanceTimersByTime(1_001);
    expect(shown(root)).toBe(false);
  });
});

describe("dismissError", () => {
  it("takes a standing error off the banner", () => {
    const root = attached();
    const toast = createToast(root);
    toast.show("Excel is not connected.", "error");
    toast.dismissError();
    expect(root.className).toBe("toast");
  });

  it("leaves a success alone: that one leaves on its own", () => {
    vi.useFakeTimers();
    const root = attached();
    const toast = createToast(root);
    toast.show("Saved");
    toast.dismissError();
    expect(shown(root)).toBe(true);
    vi.advanceTimersByTime(4_351);
    expect(shown(root)).toBe(false);
  });

  it("leaves a question open: only an answer takes it off", async () => {
    const root = attached();
    const toast = createToast(root);
    const answer = toast.ask(
      "Draw it?",
      [{ label: "Yes", value: "yes" }],
      "no",
    );
    toast.dismissError();
    expect(root.classList.contains("asking")).toBe(true);
    root.querySelector("button")!.click();
    await expect(answer).resolves.toBe("yes");
  });

  it("does nothing when the banner is already empty", () => {
    const root = attached();
    const toast = createToast(root);
    expect(() => toast.dismissError()).not.toThrow();
    expect(root.className).toBe("");
  });
});

describe("a question", () => {
  it("is not taken off by a click on its words", async () => {
    const root = attached();
    const toast = createToast(root);
    const answer = toast.ask(
      "Draw it?",
      [{ label: "Yes", value: "yes" }],
      "no",
    );
    root
      .querySelector(".toast-text")!
      .dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(shown(root)).toBe(true);
    root.querySelector("button")!.click();
    await expect(answer).resolves.toBe("yes");
  });
});
