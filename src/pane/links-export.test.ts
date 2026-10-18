// @vitest-environment jsdom
// The Links tab's four export buttons, off the tab: the markup behind them
// (the shipped taskpane.html), which adapter each press reaches, how a press is
// reported in either transport, and what a table over the cap does. Export
// selection sends a native table, a picture only by its own button, and the
// Excel adapter is mocked so no Office host is needed.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  exportActiveChart,
  exportSelection,
  exportSelectionAsTable,
  exportSelectionAsText,
} from "../excel";
import { TABLE_TOO_BIG } from "../link/model";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import {
  exportActions,
  type ExportHooks,
  type ExportTab,
} from "./links-export";

vi.mock("../excel", () => ({
  exportActiveChart: vi.fn(),
  exportSelection: vi.fn(),
  exportSelectionAsTable: vi.fn(),
  exportSelectionAsText: vi.fn(),
}));

// Stand-ins the adapter mocks never look inside: a name tells two apart, so a
// call's arguments say which key and which relay a press used.
function stand<T>(name: string): T {
  return { name } as T;
}

const PAIRED = stand<Workspace>("paired");
const RELAY = stand<RelayApi>("relay");
// What local mode's copy() hands the work: the fixed local key and a collector.
const LOCAL_KEY = stand<Workspace>("local");
const COLLECTOR = stand<RelayApi>("collector");

const BUTTONS = [
  "export-selection",
  "export-chart",
  "export-text",
  "export-picture",
];

beforeEach(() => {
  vi.resetAllMocks();
});

interface Fixture {
  actions: Map<string, () => Promise<string>>;
  hooks: ExportHooks;
}

function fixture(
  options: { mode?: "local" | "relay"; pick?: string } = {},
): Fixture {
  const chartPick = document.createElement("select");
  chartPick.append(
    new Option("Selected chart", ""),
    new Option("Revenue chart", "Revenue chart"),
  );
  chartPick.value = options.pick ?? "";
  const tab: ExportTab = {
    deps: { relay: RELAY },
    chartPick,
    transport: {
      mode: () => options.mode ?? "relay",
      setMode: async () => undefined,
      retryCopy: () => "",
      // The real copy() runs the work against the local key and a collector,
      // then reports the "ready" line when the clipboard cannot be written.
      copy: async (run, lines) => lines.ready(await run(LOCAL_KEY, COLLECTOR)),
    },
  };
  const hooks: ExportHooks = {
    workspace: vi.fn(() => PAIRED),
    refresh: vi.fn(async () => undefined),
  };
  return { actions: new Map(exportActions(tab, hooks)), hooks };
}

function press(fixed: Fixture, id: string): Promise<string> {
  const run = fixed.actions.get(id);
  if (run === undefined) throw new Error(`no export button ${id}`);
  return run();
}

describe("the export buttons' markup", () => {
  const html = readFileSync(join(process.cwd(), "taskpane.html"), "utf8");
  const pane = new DOMParser().parseFromString(html, "text/html");
  const section = pane.querySelector(
    'section[aria-labelledby="export-heading"]',
  )!;

  // What a reader sees: the markup may break a long line anywhere.
  const seen = (text: string | null): string =>
    (text ?? "").replace(/\s+/g, " ").trim();

  function words(id: string): { name: string; line: string } {
    const button = section.querySelector(`#${id}`)!;
    return {
      name: seen(button.querySelector("strong")!.textContent),
      line: seen(button.querySelector("small")!.textContent),
    };
  }

  it("ships Export selection, Export active chart, Export as text and Export as picture, in that order", () => {
    const ids = [...section.querySelectorAll("button[id^=export-]")].map(
      (button) => button.id,
    );
    expect(ids).toEqual(BUTTONS);
  });

  it("has no Export as table button any more", () => {
    expect(pane.getElementById("export-table")).toBeNull();
    expect(section.textContent).not.toContain("Export as table");
  });

  it("says a table on the first button and a picture only on its own", () => {
    expect(words("export-selection")).toEqual({
      name: "Export selection",
      line: "The selected range as an editable table that refreshes in place",
    });
    expect(words("export-picture")).toEqual({
      name: "Export as picture",
      line: "The selected range as a linked picture, for a layout a table cannot keep",
    });
  });

  it("calls the chart native shapes and leaves the text button as it was", () => {
    expect(words("export-chart")).toEqual({
      name: "Export active chart",
      line: "The selected chart as native shapes that refresh in place",
    });
    expect(words("export-text")).toEqual({
      name: "Export as text",
      line: "One cell's text as a box that refreshes in place",
    });
  });
});

describe("exportActions", () => {
  it("hands back the four buttons in the order the pane shows them", () => {
    expect([...fixture().actions.keys()]).toEqual(BUTTONS);
  });

  it("sends the selection as a native table from Export selection", async () => {
    vi.mocked(exportSelectionAsTable).mockResolvedValue({
      id: "a",
      label: "Model!B4:F5 table",
    });
    const fixed = fixture();

    const line = await press(fixed, "export-selection");

    expect(line).toBe("Sent to PowerPoint: Model!B4:F5 table");
    expect(exportSelectionAsTable).toHaveBeenCalledWith(PAIRED, RELAY);
    expect(exportSelection).not.toHaveBeenCalled();
    expect(fixed.hooks.refresh).toHaveBeenCalledTimes(1);
  });

  it("sends a picture only from Export as picture", async () => {
    vi.mocked(exportSelection).mockResolvedValue({
      id: "b",
      label: "Model!B4:F5",
    });
    const fixed = fixture();

    const line = await press(fixed, "export-picture");

    expect(line).toBe("Sent to PowerPoint: Model!B4:F5");
    expect(exportSelection).toHaveBeenCalledWith(PAIRED, RELAY);
    expect(exportSelectionAsTable).not.toHaveBeenCalled();
    expect(fixed.hooks.refresh).toHaveBeenCalledTimes(1);
  });

  it("sends one cell's text from Export as text", async () => {
    vi.mocked(exportSelectionAsText).mockResolvedValue({
      id: "c",
      label: "Model!B4 text",
    });
    const fixed = fixture();

    expect(await press(fixed, "export-text")).toBe(
      "Sent to PowerPoint: Model!B4 text",
    );
    expect(exportSelectionAsText).toHaveBeenCalledWith(PAIRED, RELAY);
    expect(exportSelection).not.toHaveBeenCalled();
    expect(exportSelectionAsTable).not.toHaveBeenCalled();
  });

  it("sends the picked chart, or the selected one when none is picked", async () => {
    vi.mocked(exportActiveChart).mockResolvedValue({
      id: "d",
      label: "Model: Revenue chart",
    });

    await press(fixture({ pick: "Revenue chart" }), "export-chart");
    await press(fixture(), "export-chart");

    expect(vi.mocked(exportActiveChart).mock.calls).toEqual([
      [PAIRED, RELAY, "Revenue chart"],
      [PAIRED, RELAY, null],
    ]);
    expect(exportSelection).not.toHaveBeenCalled();
    expect(exportSelectionAsTable).not.toHaveBeenCalled();
  });

  it("says what Excel could not describe about a chart after its label", async () => {
    vi.mocked(exportActiveChart).mockResolvedValue({
      id: "e",
      label: "Model: Revenue chart",
      note: "as a picture: 7 series; shapes draw up to 6",
    });

    expect(await press(fixture(), "export-chart")).toBe(
      "Sent to PowerPoint: Model: Revenue chart (as a picture: 7 series; shapes draw up to 6)",
    );
  });

  // The cap sentence is the adapter's, refused before anything is anchored, so
  // the press ends in the toast and the list is not read again.
  it("says a selection is too big for a native table, sends nothing and lists nothing", async () => {
    vi.mocked(exportSelectionAsTable).mockRejectedValue(
      new Error(TABLE_TOO_BIG),
    );
    const fixed = fixture();

    await expect(press(fixed, "export-selection")).rejects.toThrow(
      "Too big for a native table (60 x 20). Select less, or press Export as picture.",
    );
    expect(exportSelection).not.toHaveBeenCalled();
    expect(fixed.hooks.refresh).not.toHaveBeenCalled();
  });

  it("stops every relay press at the missing link key before it asks Excel for anything", async () => {
    const fixed = fixture();
    vi.mocked(fixed.hooks.workspace).mockImplementation(() => {
      throw new Error("Generate a link key first (Links > Link key).");
    });

    for (const id of BUTTONS) {
      await expect(press(fixed, id)).rejects.toThrow("Generate a link key");
    }
    expect(exportSelection).not.toHaveBeenCalled();
    expect(exportSelectionAsTable).not.toHaveBeenCalled();
    expect(exportSelectionAsText).not.toHaveBeenCalled();
    expect(exportActiveChart).not.toHaveBeenCalled();
  });

  it("copies through the transport in local mode, with no link key asked for", async () => {
    vi.mocked(exportSelectionAsTable).mockResolvedValue({
      id: "f",
      label: "Model!B4:F5 table",
    });
    const fixed = fixture({ mode: "local" });

    const line = await press(fixed, "export-selection");

    expect(line).toBe(
      "Model!B4:F5 table is linked and ready: press Copy for PowerPoint.",
    );
    expect(exportSelectionAsTable).toHaveBeenCalledWith(LOCAL_KEY, COLLECTOR);
    expect(fixed.hooks.workspace).not.toHaveBeenCalled();
    expect(fixed.hooks.refresh).toHaveBeenCalledTimes(1);
  });
});
