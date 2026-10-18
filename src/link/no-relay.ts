// src/link/no-relay.ts
// The relay of a build that keeps every link on the user's PC (src/link/local-only.ts): every
// call is refused in words and no request is ever made, so a flow that forgot the mode still
// sends nothing. relayFor() is where both panes get their relay. Invariant: NoRelay never
// touches the network.

import { localOnly } from "./local-only";
import { RelayClient, relayBaseUrl, type RelayApi } from "./relay";

export const NO_RELAY =
  "This pls,fix keeps your links on this computer: there is no relay to send them to.";

function refuse(): Promise<never> {
  return Promise.reject(new Error(NO_RELAY));
}

export class NoRelay implements RelayApi {
  putLink: RelayApi["putLink"] = refuse;
  getLink: RelayApi["getLink"] = refuse;
  getLinkRev: RelayApi["getLinkRev"] = refuse;
  deleteLink: RelayApi["deleteLink"] = refuse;
  status: RelayApi["status"] = refuse;
  touchLinks: RelayApi["touchLinks"] = refuse;
  fetchLinks: RelayApi["fetchLinks"] = refuse;
  postInbox: RelayApi["postInbox"] = refuse;
  listInbox: RelayApi["listInbox"] = refuse;
  deleteInbox: RelayApi["deleteInbox"] = refuse;
}

export function relayFor(documentUrl: string): RelayApi {
  return localOnly()
    ? new NoRelay()
    : new RelayClient(relayBaseUrl(documentUrl));
}
