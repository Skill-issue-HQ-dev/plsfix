// Convert to table: a picture link (registry kind "range") becomes a table link.
// The entry changes kind and the link is pushed again as a table payload under
// the same id and key, so a deck's next update replaces the picture with a native
// table at the same corner and width. Owns the per-link rule - only a picture,
// only within the table cap, the registry changed only after the push landed.

import type { RegistryEntry } from "../link/model";
import type { RelayApi } from "../link/relay";
import {
  entryOf,
  readRegistry,
  resolveSources,
  sourceOf,
  workbookName,
  writeRegistry,
  type ResolvedRange,
} from "./link-anchors";
import { exclusive } from "./link-lock";
import { pushPayload, relayFailureReason } from "./link-record";
import { renderSource } from "./link-render";

export interface ConvertSummary {
  converted: number;
  // Ticked links that are not pictures: a chart, a table or a text.
  skipped: number;
  missing: number;
  failed: number;
  // One "<label>: <reason>" per link refused or not pushed.
  failures: string[];
}

export async function convertToTable(
  ids: string[],
  relay: RelayApi,
): Promise<ConvertSummary> {
  return exclusive("convert", () => convertRegistry(ids, relay));
}

// A report, not an assertion, like a push: one link refused or unpushed is
// counted and the others go on. The registry is written back once, after the
// links that landed have changed kind.
async function convertRegistry(
  ids: string[],
  relay: RelayApi,
): Promise<ConvertSummary> {
  const workbook = await workbookName();
  return Excel.run(async (context) => {
    const registry = await readRegistry(context);
    const picked = ids.map((id) => entryOf(registry, id, "convert"));
    const pictures = picked.filter((entry) => entry.kind === "range");
    const summary: ConvertSummary = {
      converted: 0,
      skipped: picked.length - pictures.length,
      missing: 0,
      failed: 0,
      failures: [],
    };
    const sources = await resolveSources(context, pictures);
    for (const [index, entry] of pictures.entries()) {
      const resolved = sources[index] ?? null;
      if (resolved === null || resolved.kind === "chart") {
        summary.missing += 1;
        continue;
      }
      await convertOne(context, entry, resolved, workbook, relay, summary);
    }
    if (summary.converted > 0) {
      writeRegistry(context, registry);
      await context.sync();
    }
    return summary;
  });
}

// The entry changes in a copy and is handed back only once the push landed, so
// a link that fails - the cap, a source Excel cannot read, a relay that refuses -
// stays the picture link it was, at the revision the deck holds.
async function convertOne(
  context: Excel.RequestContext,
  entry: RegistryEntry,
  resolved: ResolvedRange,
  workbook: string,
  relay: RelayApi,
  summary: ConvertSummary,
): Promise<void> {
  const asTable: ResolvedRange = { ...resolved, kind: "table" };
  const next: RegistryEntry = {
    ...entry,
    kind: "table",
    label: `${entry.label} table`,
  };
  try {
    // renderTable refuses past the cap before it reads a cell.
    const render = await renderSource(context, asTable);
    const src = sourceOf(workbook, entry.anchor, asTable);
    next.rev = await pushPayload(next, src, render, relay);
    next.lastPushedAt = new Date().toISOString();
    Object.assign(entry, next);
    summary.converted += 1;
  } catch (error) {
    const reason =
      relayFailureReason(error) ??
      (error instanceof Error ? error.message : String(error));
    summary.failed += 1;
    summary.failures.push(`${entry.label}: ${reason}`);
  }
}
