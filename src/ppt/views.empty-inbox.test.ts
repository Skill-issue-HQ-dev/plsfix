// @vitest-environment jsdom
// renderInbox with nothing waiting: the line follows the link mode the pane
// passes, and stays the relay's words where no mode is given.
import { describe, expect, it } from "vitest";
import { renderInbox } from "./views";

function emptyLine(mode?: "local" | "relay"): string | undefined {
  const list = document.createElement("div");
  renderInbox(list, [], () => undefined, mode);
  return list.querySelector(".hint")?.textContent ?? undefined;
}

describe("renderInbox with nothing waiting", () => {
  it("says what to do in copy-and-paste mode", () => {
    expect(emptyLine("local")).toBe(
      "Nothing waiting. In Excel press Export, then paste here.",
    );
  });

  it("says what to do in relay mode", () => {
    expect(emptyLine("relay")).toBe(
      "Nothing waiting. Push a range or a chart from Excel.",
    );
  });

  it("keeps the relay's words where the caller names no mode", () => {
    expect(emptyLine()).toBe(
      "Nothing waiting. Push a range or a chart from Excel.",
    );
  });
});
