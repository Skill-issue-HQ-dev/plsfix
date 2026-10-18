// The custom styles' pure core, part three: the registry codec for the workbook
// setting plsfix.styles.v1 and the list edits over it. No Office.js.

import { describe, expect, it } from "vitest";
import {
  decodeRegistry,
  encodeRegistry,
  isRegistered,
  NEWER_REGISTRY,
  onlyExisting,
  registeredNames,
  REGISTRY_SETTING,
  withoutStyle,
  withStyle,
  type RegistryEntry,
} from "./custom-styles";

describe("the registry codec", () => {
  const entries: RegistryEntry[] = [
    { name: "pls,fix Title", kind: "starter" },
    { name: "Margin %", kind: "user" },
  ];

  it("lives in the setting plsfix.styles.v1", () => {
    expect(REGISTRY_SETTING).toBe("plsfix.styles.v1");
  });

  it("round-trips in list order", () => {
    const json = encodeRegistry(entries);
    expect(JSON.parse(json)).toEqual({
      v: 1,
      styles: [
        { name: "pls,fix Title", kind: "starter" },
        { name: "Margin %", kind: "user" },
      ],
    });
    expect(decodeRegistry(json)).toEqual({ ok: true, entries });
  });

  it("reads no setting, a blank one and garbage as an empty list", () => {
    for (const raw of [null, undefined, "", "   ", "{", "[]", "42", "null"]) {
      expect(decodeRegistry(raw), String(raw)).toEqual({
        ok: true,
        entries: [],
      });
    }
  });

  it("skips entries that are not a name and a kind", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        { name: "Good one", kind: "user" },
        { name: "", kind: "user" },
        { name: "No kind" },
        { name: "Odd kind", kind: "pirate" },
        "text",
        null,
        { name: 7, kind: "user" },
      ],
    });
    expect(decodeRegistry(raw)).toEqual({
      ok: true,
      entries: [{ name: "Good one", kind: "user" }],
    });
  });

  it("keeps the first of a name given twice, whatever its case", () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [
        { name: "Margin %", kind: "user" },
        { name: "MARGIN %", kind: "user" },
      ],
    });
    expect(decodeRegistry(raw)).toEqual({
      ok: true,
      entries: [{ name: "Margin %", kind: "user" }],
    });
  });

  it("reads a newer version as: update pls,fix", () => {
    const raw = JSON.stringify({ v: 2, styles: [{ name: "A", kind: "user" }] });
    expect(decodeRegistry(raw)).toEqual({ ok: false, message: NEWER_REGISTRY });
    expect(NEWER_REGISTRY).toMatch(/newer pls,fix\. Update the add-in\.$/);
  });

  it("reads a missing version as the first one", () => {
    const raw = JSON.stringify({ styles: [{ name: "A", kind: "user" }] });
    expect(decodeRegistry(raw)).toEqual({
      ok: true,
      entries: [{ name: "A", kind: "user" }],
    });
  });

  it("lists the names of any version, for the scrubber's exemption", () => {
    const raw = JSON.stringify({
      v: 9,
      styles: [{ name: "Future", kind: "user" }, { name: 5 }, "x"],
    });
    expect(registeredNames(raw)).toEqual(["Future"]);
    expect(registeredNames(null)).toEqual([]);
    expect(registeredNames("{")).toEqual([]);
  });
});

describe("registry list edits", () => {
  const starter = (name: string): RegistryEntry => ({ name, kind: "starter" });
  const user = (name: string): RegistryEntry => ({ name, kind: "user" });

  it("puts the newest user style last", () => {
    const next = withStyle([user("A")], user("B"));
    expect(next.map((entry) => entry.name)).toEqual(["A", "B"]);
  });

  it("keeps the starters first, in the starter order, whatever order they arrive in", () => {
    let list: RegistryEntry[] = [user("Mine")];
    list = withStyle(list, starter("pls,fix Header"));
    list = withStyle(list, starter("pls,fix Title"));
    expect(list.map((entry) => entry.name)).toEqual([
      "pls,fix Title",
      "pls,fix Header",
      "Mine",
    ]);
  });

  it("replaces an entry of the same name in place instead of repeating it", () => {
    const next = withStyle([user("A"), user("B")], user("a"));
    expect(next).toEqual([user("A"), user("B")]);
  });

  it("removes by name, whatever the case, and leaves the rest alone", () => {
    expect(withoutStyle([user("A"), user("B")], "b")).toEqual([user("A")]);
    expect(withoutStyle([user("A")], "Z")).toEqual([user("A")]);
  });

  it("answers whether a name is registered, whatever its case", () => {
    expect(isRegistered([user("Margin %")], "margin %")).toBe(true);
    expect(isRegistered([user("Margin %")], "Other")).toBe(false);
  });

  it("keeps only the styles the workbook still has", () => {
    const list = [user("A"), starter("pls,fix Title"), user("Gone")];
    expect(onlyExisting(list, ["Normal", "a", "PLS,FIX TITLE"])).toEqual([
      user("A"),
      starter("pls,fix Title"),
    ]);
  });
});
