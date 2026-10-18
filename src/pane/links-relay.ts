// src/pane/links-relay.ts
// What the Excel Links tab asks of the relay, one decision per flow, so
// links-tab.ts (already past its length cap) only calls them. Owns: the relay
// half of Push selected / Push all, and which relay Remove link talks to.
// Invariants: a relay push also posts the Inbox row of every link it pushes
// (only when this computer holds the key that seals the row); in local mode
// no flow here reaches the real relay (design 2026-09-23 s.6).

import { pushLinks, type PushSummary } from "../excel";
import { LocalCollector } from "../link/local-collector";
import type { RelayApi } from "../link/relay";
import type { LinkTransport } from "../link/transport-setting";
import type { Workspace } from "../link/workspace";

// Push selected / Push all through the relay. With a link key each pushed link
// also posts the row an export posts (the same sealed envelope, via announce);
// with none - a computer that never made one, or whose key could not be read -
// the revisions still go out and the Inbox is left alone, as it always was.
export function pushThroughRelay(
  ids: string[] | "all",
  relay: RelayApi,
  ws: Workspace | null,
): Promise<PushSummary> {
  return ws === null
    ? pushLinks(ids, relay)
    : pushLinks(ids, relay, { announce: ws });
}

// The relay Remove link revokes a link's copy on. Relay mode: the real one,
// whose refusal keeps the link in the workbook so the revoke can be retried.
// Local mode: no relay copy exists to revoke and no request may leave the
// computer, so a LocalCollector (its deleteLink does nothing) stands in and
// the deck keeps its object.
export function removalRelay(mode: LinkTransport, relay: RelayApi): RelayApi {
  return mode === "local" ? new LocalCollector() : relay;
}
