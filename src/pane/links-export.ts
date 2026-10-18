// The Links tab's four export buttons: what each asks Excel for and how a press
// is reported. Export selection sends a native table; Export as picture is the
// only way a range becomes a picture. Invariants: an adapter is looked up when a
// button is pressed, never at load, and local copy() starts before any await.

import {
  exportActiveChart,
  exportSelection as exportSelectionAsPicture,
  exportSelectionAsTable,
  exportSelectionAsText,
} from "../excel";
import type { RelayApi } from "../link/relay";
import type { Workspace } from "../link/workspace";
import { exportCopyLines, type LinksTransport } from "./links-transport";

// The slice of the tab an export press reads. The transport is read when a
// button is pressed: boot() swaps the real one in after the tab is wired.
export interface ExportTab {
  deps: { relay: RelayApi };
  transport: LinksTransport;
  chartPick: HTMLSelectElement;
}

export interface ExportHooks {
  /** The paired link key, or the sentence for why there is none. */
  workspace(): Workspace;
  /** The link list, read again once an export has landed. */
  refresh(): Promise<void>;
}

export type ExportAction = [id: string, run: () => Promise<string>];

type RangeExport = "table" | "picture" | "text";

interface Sent {
  label: string;
  note?: string;
}

type Send = (ws: Workspace, relay: RelayApi) => Promise<Sent>;

/** The four buttons, in the order the pane shows them, and what each sends. */
export function exportActions(
  tab: ExportTab,
  hooks: ExportHooks,
): ExportAction[] {
  return [
    ["export-selection", () => exportRange(tab, hooks, "table")],
    ["export-chart", () => exportChart(tab, hooks)],
    ["export-text", () => exportRange(tab, hooks, "text")],
    ["export-picture", () => exportRange(tab, hooks, "picture")],
  ];
}

function sender(kind: RangeExport): Send {
  if (kind === "table") return exportSelectionAsTable;
  return kind === "text" ? exportSelectionAsText : exportSelectionAsPicture;
}

function exportRange(
  tab: ExportTab,
  hooks: ExportHooks,
  kind: RangeExport,
): Promise<string> {
  return send(tab, hooks, (ws, relay) => sender(kind)(ws, relay));
}

// With nothing selected the picker names the chart; its value is read now,
// before anything is awaited.
function exportChart(tab: ExportTab, hooks: ExportHooks): Promise<string> {
  const pick = tab.chartPick.value || null;
  return send(tab, hooks, (ws, relay) => exportActiveChart(ws, relay, pick));
}

// Local mode copies the export for PowerPoint instead of pushing it; relay
// mode needs the paired key and pushes. Either way the list is read again.
async function send(
  tab: ExportTab,
  hooks: ExportHooks,
  run: Send,
): Promise<string> {
  if (tab.transport.mode() === "local") {
    return tab.transport.copy(async (ws, relay) => {
      const result = await run(ws, relay);
      await hooks.refresh();
      return result;
    }, exportCopyLines());
  }
  const result = await run(hooks.workspace(), tab.deps.relay);
  await hooks.refresh();
  return sentLine(result);
}

// The toast: the label, and for a chart Excel could not describe the sentence
// PowerPoint will repeat beside the picture.
function sentLine(result: Sent): string {
  const note = result.note === undefined ? "" : ` (${result.note})`;
  return `Sent to PowerPoint: ${result.label}${note}`;
}
