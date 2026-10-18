// Toast: the one status banner shared by every pane. A success stays long
// enough to read (src/ui/toast-timing.ts); an error has no timer and stays until
// the next pane action takes it off (dismissError, which the guard calls) or a
// click on the toast does; an error with details renders a "Copy details"
// button so a failure can be pasted into a bug report; a question (ask) holds
// the banner with one button per answer until one is pressed, puts focus on the
// first answer and takes Escape as the fallback. No Office.js here - this
// module only ever touches the DOM node it is given.
// Invariant: a question is never left unanswered - whatever takes the banner
// from it answers it with its fallback.

import { COPY_FAILED_MESSAGE, copyText } from "./clipboard";
import { successStayMs } from "./toast-timing";

export type ToastKind = "success" | "error";

export interface Toast {
  // `stayMs` keeps this one message up for that long instead of the toast's
  // own rule (a success by its length, an error until it is acted on): a draw
  // reports each chunk through it, and a chunk can take far longer than a
  // sentence needs to be read.
  show(
    message: string,
    kind?: ToastKind,
    details?: string,
    stayMs?: number,
  ): void;
}

// One answer to a question: what the button says and what it gives back.
export interface Choice<T> {
  label: string;
  value: T;
}

export interface AskingToast extends Toast {
  // The banner shows `message` with one button per choice and stays up until
  // one is pressed, then answers with its value. Anything else that takes the
  // banner first (another message, a newer question) answers `dismissed`.
  ask<T>(
    message: string,
    choices: readonly Choice<T>[],
    dismissed: T,
  ): Promise<T>;
}

// The banner every pane builds: the shared toast plus the one thing the guard
// asks of it between actions.
export interface PaneToast extends AskingToast {
  // Takes a standing error (one with no timer) off the banner. The guard calls
  // it as each pane action starts, so the last failure never outlives the next
  // action. A success leaves on its own and a question waits for its answer:
  // neither is touched.
  dismissError(): void;
}

const DISMISS_HINT = "Click to dismiss";

// What the banner's functions share: the node they paint, the clock on its
// message and the two things a message can be waiting for.
interface Banner {
  container: HTMLElement;
  // A fixed time every success stays, instead of successStayMs(message).
  successMs: number | undefined;
  timer: number | undefined;
  // True while the banner shows an error with no timer: it waits for the next
  // pane action (dismissStanding) or a click on the toast.
  standing: boolean;
  // The question on the banner, when one is open: whatever takes the banner
  // from it calls dismiss, which answers it with its fallback.
  open: { dismiss: () => void } | null;
}

// A drag over the toast's words selects them to copy by hand: the click that
// ends it is not a dismissal.
function holdsSelection(container: HTMLElement): boolean {
  const selection = window.getSelection();
  return (
    selection !== null &&
    !selection.isCollapsed &&
    container.contains(selection.anchorNode)
  );
}

function toastLine(message: string): HTMLSpanElement {
  const text = document.createElement("span");
  text.className = "toast-text";
  text.textContent = message;
  return text;
}

function toastButton(label: string, onPress: () => void): HTMLButtonElement {
  const control = document.createElement("button");
  control.type = "button";
  control.className = "toast-copy";
  control.textContent = label;
  control.addEventListener("click", onPress);
  return control;
}

// How long a message stays, or undefined for an error with no time of its own,
// which waits to be acted on. A caller's own time wins; a success stays by its
// length, or for the fixed time the toast was built with.
function stayFor(
  banner: Banner,
  message: string,
  kind: ToastKind,
  stayMs: number | undefined,
): number | undefined {
  if (stayMs !== undefined) return stayMs;
  if (kind === "error") return undefined;
  return banner.successMs ?? successStayMs(message);
}

// Takes the message off the banner and stops its clock.
function hide(banner: Banner): void {
  window.clearTimeout(banner.timer);
  banner.standing = false;
  banner.container.className = "toast";
  banner.container.removeAttribute("title");
}

// Clears the way for a new message or question: an open question is answered
// with its fallback and the clock on the last message stops.
function makeRoom(banner: Banner): void {
  banner.open?.dismiss();
  window.clearTimeout(banner.timer);
}

// The "Copy details" button of an error: copies, then says whether it worked.
function copyDetailsButton(
  details: string,
  say: (message: string, kind: ToastKind) => void,
): HTMLButtonElement {
  return toastButton("Copy details", () => {
    void copyText(details).then((copied) => {
      say(
        copied ? "Details copied." : COPY_FAILED_MESSAGE,
        copied ? "success" : "error",
      );
    });
  });
}

// Puts a message on the banner, in place of whatever was there, and starts the
// clock it has, if any.
function showMessage(
  banner: Banner,
  message: string,
  kind: ToastKind = "success",
  details?: string,
  stayMs?: number,
): void {
  const { container } = banner;
  makeRoom(banner);
  container.replaceChildren(toastLine(message));

  if (details) {
    container.append(
      copyDetailsButton(details, (text, tone) => {
        showMessage(banner, text, tone);
      }),
    );
  }

  container.className = `toast visible ${kind}`;
  const stay = stayFor(banner, message, kind, stayMs);
  banner.standing = stay === undefined;
  if (stay === undefined) container.title = DISMISS_HINT;
  else container.removeAttribute("title");
  if (stay !== undefined) {
    banner.timer = window.setTimeout(() => {
      hide(banner);
    }, stay);
  }
}

// Takes a standing error off the banner; a success leaves on its own and a
// question waits for its answer.
function dismissStanding(banner: Banner): void {
  if (banner.standing) hide(banner);
}

// A click on the toast takes it off, unless it is a question (its buttons
// answer), a button of its own (Copy details copies) or the end of a drag
// over its words.
function dismissOnClick(banner: Banner): void {
  const { container } = banner;
  container.addEventListener("click", (event) => {
    if (banner.open !== null || !container.classList.contains("visible")) {
      return;
    }
    const target = event.target;
    if (target instanceof Element && target.closest("button") !== null) return;
    if (holdsSelection(container)) return;
    hide(banner);
  });
}

// The question's words and one button per answer, with focus on the first.
// "asking" lets the buttons wrap under the question in a narrow pane.
function renderQuestion<T>(
  container: HTMLElement,
  message: string,
  choices: readonly Choice<T>[],
  answer: (value: T) => void,
): void {
  container.replaceChildren(
    toastLine(message),
    ...choices.map((choice) =>
      toastButton(choice.label, () => {
        answer(choice.value);
      }),
    ),
  );
  container.className = "toast visible success asking";
  container.querySelector("button")?.focus();
}

// One question on the banner. It ends exactly once, by an answer, an Escape or
// whatever takes the banner, and its key listener goes with it.
function askQuestion<T>(
  banner: Banner,
  message: string,
  choices: readonly Choice<T>[],
  dismissed: T,
): Promise<T> {
  makeRoom(banner);
  // A question is not an error waiting to be acted on.
  banner.standing = false;
  banner.container.removeAttribute("title");
  return new Promise<T>((resolve) => {
    function finish(value: T): void {
      document.removeEventListener("keydown", onKey);
      resolve(value);
    }
    const mine = {
      dismiss: () => {
        banner.open = null;
        finish(dismissed);
      },
    };
    // A button answers once, and only while its own question is the open
    // one: a stale button never answers a newer question.
    function answer(value: T): void {
      if (banner.open !== mine) return;
      banner.open = null;
      banner.container.className = "toast";
      banner.container.replaceChildren();
      finish(value);
    }
    function onKey(event: KeyboardEvent): void {
      if (event.key === "Escape") answer(dismissed);
    }
    banner.open = mine;
    document.addEventListener("keydown", onKey);
    renderQuestion(banner.container, message, choices, answer);
  });
}

// `successMs`, when given, is how long every success stays instead of
// successStayMs(message): the UX gate and the tests pin it.
export function createToast(
  container: HTMLElement,
  successMs?: number,
): PaneToast {
  const banner: Banner = {
    container,
    successMs,
    timer: undefined,
    standing: false,
    open: null,
  };
  dismissOnClick(banner);
  return {
    show: (message, kind, details, stayMs) => {
      showMessage(banner, message, kind, details, stayMs);
    },
    ask: (message, choices, dismissed) =>
      askQuestion(banner, message, choices, dismissed),
    dismissError: () => {
      dismissStanding(banner);
    },
  };
}
