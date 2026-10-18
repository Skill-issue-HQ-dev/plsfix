// The Links tab's Convert to table: what a press asks Excel for and how it is
// told - the counts as a toast, every refusal as a sentence. Local mode copies
// the table for PowerPoint instead of pushing it. Invariants: the list is read
// again once a press has run, and in local mode copy() starts before any await
// (WebKit only allows the clipboard write inside the press).

import { convertToTable, type ConvertSummary } from "../excel";
import type { RelayApi } from "../link/relay";
import type { LinksTransport } from "./links-transport";

// The slice of the tab a press reads. The transport is read when the button is
// pressed: boot() swaps the real one in after the tab is wired.
export interface ConvertTab {
  deps: { relay: RelayApi };
  transport: LinksTransport;
}

export const NOT_A_PICTURE =
  "Only a picture link can become a table: tick one made with Export as picture.";

export function convertLine(summary: ConvertSummary): string {
  const parts = [
    [
      summary.converted,
      summary.converted === 1 ? "converted to a table" : "converted to tables",
    ],
    [summary.skipped, "not a picture"],
    [summary.missing, "missing"],
    [summary.failed, "failed"],
  ] as const;
  const text = parts
    .filter(([count]) => count > 0)
    .map(([count, word]) => `${String(count)} ${word}`)
    .join(", ");
  return text || "Nothing to convert";
}

// A copy's two lines: the toast once the clipboard took it, and the one that
// asks for the Copy for PowerPoint button when it did not.
function convertCopyLines(): {
  copied: (summary: ConvertSummary) => string;
  ready: (summary: ConvertSummary) => string;
} {
  return {
    copied: (summary) =>
      `${convertLine(summary)}. Paste it in PowerPoint: pls,fix, Inbox tab.`,
    ready: (summary) => `${convertLine(summary)}: press Copy for PowerPoint.`,
  };
}

export interface ConvertHooks {
  /** The ticked links' ids, or the sentence for why there are none. */
  selection(): string[];
  /** The link list, read again once a press has run. */
  refresh(): Promise<void>;
}

/** The one button, as the [id, run] pair links-tab.ts wires. */
export function convertActions(
  tab: ConvertTab,
  hooks: ConvertHooks,
): [id: string, run: () => Promise<string>][] {
  return [
    [
      "convert-to-table",
      () => convertSelected(tab, hooks.selection(), hooks.refresh),
    ],
  ];
}

// The ticked links, converted. A link that could not be converted says why in
// the thrown sentence, after the ones that could have been; ticking nothing
// that is a picture says that instead.
export async function convertSelected(
  tab: ConvertTab,
  ids: string[],
  refresh: () => Promise<void>,
): Promise<string> {
  let summary: ConvertSummary | undefined;
  const run = async (relay: RelayApi): Promise<ConvertSummary> => {
    summary = await convertToTable(ids, relay);
    await refresh();
    return summary;
  };
  const line =
    tab.transport.mode() === "local"
      ? await tab.transport.copy((_ws, relay) => run(relay), convertCopyLines())
      : convertLine(await run(tab.deps.relay));
  if (summary === undefined) return line;
  if (summary.failures.length > 0) {
    throw new Error(`${convertLine(summary)}. ${summary.failures.join(" ")}`);
  }
  const none = summary.converted === 0 && summary.missing === 0;
  if (none && summary.skipped > 0) throw new Error(NOT_A_PICTURE);
  return line;
}
