// emptyInboxLine and updateAllReceipt: what the PowerPoint pane says that
// depends on how links reach this computer.
import { describe, expect, it } from "vitest";
import { emptyInboxLine, updateAllReceipt } from "./mode-text";

describe("emptyInboxLine", () => {
  it("relay: tells the user to push from Excel, as it always did", () => {
    expect(emptyInboxLine("relay")).toBe(
      "Nothing waiting. Push a range or a chart from Excel.",
    );
  });

  it("local: tells the user to press Export in Excel, then paste here", () => {
    expect(emptyInboxLine("local")).toBe(
      "Nothing waiting. In Excel press Export, then paste here.",
    );
  });
});

describe("updateAllReceipt", () => {
  it("relay: leaves the receipt as it was", () => {
    expect(updateAllReceipt("2 updated, 1 up to date", "relay")).toBe(
      "2 updated, 1 up to date",
    );
  });

  it("local: adds how to update from Excel", () => {
    expect(updateAllReceipt("2 updated, 1 up to date", "local")).toBe(
      "2 updated, 1 up to date. To update from Excel, copy there and paste on the Inbox tab.",
    );
  });

  it("local: does not double a full stop the receipt already ends with", () => {
    expect(updateAllReceipt("No links found.", "local")).toBe(
      "No links found. To update from Excel, copy there and paste on the Inbox tab.",
    );
  });
});
