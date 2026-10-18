// Shared scaffolding of the chart-export suites that need only an exported chart
// and the payload it put on the relay: a booted strict Excel fake with a relay
// and a workspace, one chart added and exported by name, the picture payload
// read back. Owns no assertions; each suite states its own.

import { vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  decodePayload,
  REGISTRY_SETTING,
  type PicturePayload,
} from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";

export interface ChartBoot {
  links: typeof LinksModule;
  helpers: FakeHelpers;
  relay: FakeRelay;
  ws: Workspace;
  made: number;
}

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

// A fresh host with a Model and a Data sheet, a relay and a workspace; `web`
// makes it Excel for the web (the fake's web-shaped host).
export async function bootChartHost(
  options: { web?: boolean } = {},
): Promise<ChartBoot> {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"], ...options });
  host.workbook.fileUrl = "/Volumes/Work/Models/Model_v4.xlsx";
  return {
    links: await import("../src/excel/links"),
    helpers: host.helpers,
    relay: new FakeRelay(),
    ws: await createWorkspace(memoryStore()),
    made: 0,
  };
}

// The picture payload an export put on the relay, opened with its own token.
export async function pictureOf(
  boot: ChartBoot,
  id: string,
): Promise<PicturePayload> {
  const registry = JSON.parse(String(boot.helpers.setting(REGISTRY_SETTING)));
  const entry = registry.links.find((one: { id: string }) => one.id === id);
  const keys = await deriveLinkKeys(entry.token);
  const payload = decodePayload(
    await open(keys.enc, id, boot.relay.links.get(id)!.blob),
  );
  if (payload.kind !== "picture") throw new Error(`${id}: not a picture`);
  return payload;
}

// One chart added to the Model sheet and exported as a link. A null chart type
// is the web's answer for a combo: the chart is then named and titled "Mixed".
export async function exportChart(
  boot: ChartBoot,
  chartType: string | null,
  series: object[],
  extra: object = {},
) {
  boot.made += 1;
  const name = chartType ?? "Mixed";
  const chart = boot.helpers.addChart("Model", {
    name: `${name} ${String(boot.made)}`,
    chartType,
    title: name,
    series,
    ...extra,
  });
  boot.helpers.setActiveChart(chart);
  const result = await boot.links.exportActiveChart(boot.ws, boot.relay);
  return { result, payload: await pictureOf(boot, result.id) };
}
