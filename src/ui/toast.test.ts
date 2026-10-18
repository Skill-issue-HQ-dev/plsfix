// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { createToast } from "./toast";

// A test that turned the fake clock on must not leave it on: a hanging await
// would then never reach vitest's own timeout.
afterEach(() => {
  vi.useRealTimers();
});

describe("toast", () => {
  it("shows the message, then hides", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    toast.show("Saved");
    expect(root.className).toContain("visible");
    expect(root.textContent).toContain("Saved");
    vi.advanceTimersByTime(101);
    expect(root.className).toBe("toast");
  });
  it("offers a copy button only when details exist", async () => {
    const root = document.createElement("div");
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root, 100).show("Failed", "error", "stack...");
    const button = root.querySelector("button")!;
    button.click();
    expect(write).toHaveBeenCalledWith("stack...");
    createToast(root, 100).show("Plain");
    expect(root.querySelector("button")).toBeNull();
  });

  it("reports 'Details copied.' once the copy actually lands", async () => {
    const root = document.createElement("div");
    const write = vi.fn(() => Promise.resolve());
    Object.assign(navigator, { clipboard: { writeText: write } });
    createToast(root, 100).show("Failed", "error", "stack...");

    root.querySelector("button")!.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(root.textContent).toContain("Details copied.");
    expect(root.className).toContain("success");
  });

  it("reports the copy failed when neither clipboard path works", async () => {
    const root = document.createElement("div");
    Object.assign(navigator, { clipboard: undefined });
    Object.assign(document, { execCommand: vi.fn(() => false) });
    createToast(root, 100).show("Failed", "error", "stack...");

    root.querySelector("button")!.click();
    await Promise.resolve();
    await Promise.resolve();

    expect(root.textContent).toContain(
      "Copy failed: select the text and copy it by hand.",
    );
    expect(root.className).toContain("error");
  });
});

// A draw reports every chunk through the toast, and a chunk can take longer
// than the default three seconds: the caller says how long its line stays.
describe("a message that stays", () => {
  it("is hidden after the time the caller gave, not the default", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    toast.show("Drawing: 12 of 92 shapes.", "success", undefined, 5_000);
    vi.advanceTimersByTime(4_999);
    expect(root.className).toContain("visible");
    vi.advanceTimersByTime(2);
    expect(root.className).toBe("toast");
    vi.useRealTimers();
  });

  it("goes back to the default for the next plain message", () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    toast.show("Drawing: 12 of 92 shapes.", "success", undefined, 5_000);
    toast.show("Inserted Model!B4:F12.");
    vi.advanceTimersByTime(101);
    expect(root.className).toBe("toast");
    vi.useRealTimers();
  });
});

// The one question the pane asks: a long draw, "Draw it" or "Insert as
// picture". It stays until answered, and nothing else may leave it dangling.
describe("a question", () => {
  const CHOICES = [
    { label: "Draw it (about 3 min)", value: "draw" },
    { label: "Insert as picture", value: "picture" },
  ] as const;

  function buttons(root: HTMLElement): HTMLButtonElement[] {
    return [...root.querySelectorAll<HTMLButtonElement>("button")];
  }

  it("shows the question with one button per choice and answers with the pressed one", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const answer = toast.ask(
      "Model: B4:F12: 125 shapes will take about 3 min.",
      CHOICES,
      "cancel",
    );

    expect(root.className).toContain("visible");
    expect(root.querySelector(".toast-text")!.textContent).toBe(
      "Model: B4:F12: 125 shapes will take about 3 min.",
    );
    expect(buttons(root).map((button) => button.textContent)).toEqual([
      "Draw it (about 3 min)",
      "Insert as picture",
    ]);
    // The same bordered button the toast's own "Copy details" is, the one
    // control the fixed bar lets through to the pane under it.
    expect(
      buttons(root).every((button) => button.className === "toast-copy"),
    ).toBe(true);
    expect(buttons(root).every((button) => button.type === "button")).toBe(
      true,
    );

    buttons(root)[1]!.click();
    await expect(answer).resolves.toBe("picture");
    expect(root.className).toBe("toast");
    expect(buttons(root)).toHaveLength(0);
  });

  it("marks the banner as asking, so a narrow pane wraps the answers under the question", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const answer = toast.ask("Draw?", CHOICES, "cancel");
    expect(root.classList.contains("asking")).toBe(true);

    toast.show("Next.");
    await expect(answer).resolves.toBe("cancel");
    expect(root.classList.contains("asking")).toBe(false);
  });

  it("stays up past the toast's own hide time", async () => {
    vi.useFakeTimers();
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const answer = toast.ask("Draw?", CHOICES, "cancel");
    vi.advanceTimersByTime(60_000);
    expect(root.className).toContain("visible");
    expect(buttons(root)).toHaveLength(2);
    buttons(root)[0]!.click();
    await expect(answer).resolves.toBe("draw");
    vi.useRealTimers();
  });

  it("keeps the first answer when a button is pressed twice", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const answer = toast.ask("Draw?", CHOICES, "cancel");
    const [draw, picture] = buttons(root);
    draw!.click();
    picture!.click();
    await expect(answer).resolves.toBe("draw");
  });

  it("is dismissed, not left hanging, when another message takes the toast", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const answer = toast.ask("Draw?", CHOICES, "cancel");

    toast.show("Wait for the last action to finish.", "error");

    await expect(answer).resolves.toBe("cancel");
    expect(root.textContent).toContain("Wait for the last action to finish.");
    expect(buttons(root)).toHaveLength(0);
  });

  it("never lets a button of a dismissed question answer a newer one", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const first = toast.ask("First?", CHOICES, "cancel");
    const stale = buttons(root)[0]!;
    const second = toast.ask("Second?", CHOICES, "cancel");

    stale.click();

    expect(root.querySelector(".toast-text")!.textContent).toBe("Second?");
    buttons(root)[1]!.click();
    await expect(second).resolves.toBe("picture");
    await expect(first).resolves.toBe("cancel");
  });

  it("is dismissed by a newer question, which then owns the toast", async () => {
    const root = document.createElement("div");
    const toast = createToast(root, 100);
    const first = toast.ask("First?", CHOICES, "cancel");
    const second = toast.ask("Second?", CHOICES, "cancel");

    await expect(first).resolves.toBe("cancel");
    expect(root.querySelector(".toast-text")!.textContent).toBe("Second?");
    buttons(root)[0]!.click();
    await expect(second).resolves.toBe("draw");
  });

  describe("from the keyboard", () => {
    function attached(): HTMLElement {
      const root = document.createElement("div");
      document.body.append(root);
      return root;
    }

    afterEach(() => {
      document.body.replaceChildren();
    });

    function press(key: string, on: EventTarget = document): void {
      on.dispatchEvent(
        new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true }),
      );
    }

    it("puts focus on the first answer when the question opens", () => {
      const root = attached();
      const toast = createToast(root, 100);
      void toast.ask("Draw?", CHOICES, "cancel");

      expect(document.activeElement).toBe(buttons(root)[0]);
    });

    it("answers the dismissed fallback on Escape, wherever the focus is", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const answer = toast.ask("Draw?", CHOICES, "cancel");

      press("Escape");

      await expect(answer).resolves.toBe("cancel");
      expect(root.className).toBe("toast");
      expect(buttons(root)).toHaveLength(0);
    });

    it("answers on Escape pressed on a button too (it bubbles up to the document)", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const answer = toast.ask("Draw?", CHOICES, "cancel");

      press("Escape", buttons(root)[0]!);

      await expect(answer).resolves.toBe("cancel");
    });

    it("ignores every other key", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      let settled = false;
      void toast.ask("Draw?", CHOICES, "cancel").then(() => (settled = true));

      press("Enter");
      press("a");
      press("Tab");
      await Promise.resolve();

      expect(settled).toBe(false);
      expect(buttons(root)).toHaveLength(2);
    });

    it("lets the Enter key press the focused button like any button", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const answer = toast.ask("Draw?", CHOICES, "cancel");

      // A real Enter on a focused button is a click; the question is answered
      // with that button's value.
      buttons(root)[0]!.click();

      await expect(answer).resolves.toBe("draw");
    });

    it("stops listening once the question is answered: a later Escape touches nothing", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const answer = toast.ask("Draw?", CHOICES, "cancel");
      buttons(root)[1]!.click();
      await expect(answer).resolves.toBe("picture");

      toast.show("Inserted.");
      press("Escape");

      expect(root.textContent).toContain("Inserted.");
      expect(root.className).toContain("visible");
    });

    it("stops listening once another message took the toast", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const answer = toast.ask("Draw?", CHOICES, "cancel");
      toast.show("Cancelled.");
      await expect(answer).resolves.toBe("cancel");

      press("Escape");

      expect(root.textContent).toContain("Cancelled.");
      expect(root.className).toContain("visible");
    });

    it("lets only the newest of two questions answer an Escape", async () => {
      const root = attached();
      const toast = createToast(root, 100);
      const first = toast.ask("First?", CHOICES, "cancel");
      const second = toast.ask("Second?", CHOICES, "dismissed-second");

      press("Escape");

      await expect(first).resolves.toBe("cancel");
      await expect(second).resolves.toBe("dismissed-second");
      expect(root.className).toBe("toast");
    });
  });
});
