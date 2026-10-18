// Where a workbook remembers which of its cell styles pls,fix made: the setting
// plsfix.styles.v1, read in the same batch as whatever else a flow loads and
// written with the flow's own write. The list itself is src/custom-styles.ts.

import {
  decodeRegistry,
  encodeRegistry,
  REGISTRY_SETTING,
  type RegistryEntry,
  type RegistryRead,
} from "../custom-styles";

/** Queues the read; the caller's next sync fills it. */
export function queueRegistryRead(
  context: Excel.RequestContext,
): Excel.Setting {
  const setting =
    context.workbook.settings.getItemOrNullObject(REGISTRY_SETTING);
  setting.load("isNullObject,value");
  return setting;
}

/** The stored JSON of a read that has synced, or null when none was saved. */
export function registryJson(setting: Excel.Setting): string | null {
  if (setting.isNullObject || typeof setting.value !== "string") return null;
  return setting.value;
}

export function registryOf(setting: Excel.Setting): RegistryRead {
  return decodeRegistry(registryJson(setting));
}

/**
 * The list a read has synced, in list order. A list a newer pls,fix wrote throws
 * its sentence: this build reads it as nothing and must not write over it.
 */
export function registryEntries(setting: Excel.Setting): RegistryEntry[] {
  const read = registryOf(setting);
  if (!read.ok) throw new Error(read.message);
  return read.entries;
}

/** Queues the new list; add() replaces the value under the same key. */
export function queueRegistryWrite(
  context: Excel.RequestContext,
  entries: readonly RegistryEntry[],
): void {
  context.workbook.settings.add(REGISTRY_SETTING, encodeRegistry(entries));
}
