// @vitest-environment jsdom
import { describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../link/model";
import {
  renderCandidates,
  renderInbox,
  renderLinkRows,
  renderSlideOptions,
  statusDetail,
  statusLabel,
  type LinkRowView,
} from "./views";

function waiting(workbook: string, label = "Model!B4:F12"): InboxItem {
  return {
    id: workbook,
    token: "t",
    kind: "range",
    label,
    src: {
      workbook,
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_aaaaaaaa",
    },
    createdAt: new Date().toISOString(),
  };
}

function linkRow(
  overrides: Partial<LinkRowView> & { key: string },
): LinkRowView {
  return {
    slide: 1,
    label: "Model!A1",
    source: "Model.xlsx",
    kind: "range",
    status: "current",
    pushedAt: null,
    selected: false,
    ...overrides,
  };
}

function rowKind(tr: Element): "folder" | "row" {
  return tr.classList.contains("link-folder") ? "folder" : "row";
}

describe("renderLinkRows", () => {
  it("renders one row per link with a checkbox and a status badge", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    const onToggle = vi.fn();
    renderLinkRows(
      body,
      [
        {
          key: "s1/sh1",
          slide: 3,
          label: "Model!B4:F12",
          source: "Model_v4.xlsx",
          kind: "range",
          status: "updateAvailable",
          pushedAt: null,
          selected: false,
        },
      ],
      onToggle,
    );
    expect(body.querySelectorAll("tr[data-key]")).toHaveLength(1);
    expect(body.textContent).toContain("Update available");
    // The kind rides in the meta line under the label: a table link and a
    // picture of the same cells carry the same text. A picture is called what
    // it is, the user's own choice, never the registry's word "range".
    expect(body.querySelector(".link-meta")!.textContent).toBe(
      "Model_v4.xlsx · Picture (your choice)",
    );
    const badge = body.querySelector(".link-status .badge") as HTMLElement;
    expect(badge.title).toBe("Update available");
    (body.querySelector("input[type=checkbox]") as HTMLInputElement).click();
    expect(onToggle).toHaveBeenCalledWith("s1/sh1", true);
  });

  // The same words the Excel list uses: a picture the user chose, then the
  // other three kinds in one capitalised word each.
  it("names every kind in the meta line, in the Excel list's words", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [
        linkRow({ key: "p", kind: "range" }),
        linkRow({ key: "c", kind: "chart" }),
        linkRow({ key: "t", kind: "table" }),
        linkRow({ key: "x", kind: "text" }),
      ],
      () => undefined,
    );
    const metas = [...body.querySelectorAll(".link-meta")].map(
      (meta) => meta.textContent,
    );
    expect(metas).toEqual([
      "Model.xlsx · Picture (your choice)",
      "Model.xlsx · Chart",
      "Model.xlsx · Table",
      "Model.xlsx · Text",
    ]);
  });

  // A chart the user kept as a picture is a picture on the slide: its row says
  // so in the words an explicit picture export uses, and a chart that was
  // drawn as shapes still reads Chart.
  it("reads Picture (your choice) for a chart kept as a picture, and Chart for one drawn", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [
        linkRow({ key: "kept", kind: "chart", choice: "picture" }),
        linkRow({ key: "drawn", kind: "chart" }),
      ],
      () => undefined,
    );
    const metas = [...body.querySelectorAll(".link-meta")].map(
      (meta) => meta.textContent,
    );
    expect(metas).toEqual([
      "Model.xlsx · Picture (your choice)",
      "Model.xlsx · Chart",
    ]);
    expect(
      body.querySelector<HTMLElement>("tr[data-key=kept] td.link-label")!.title,
    ).toBe("Model!A1 · Model.xlsx · Picture (your choice)");
  });

  it("groups rows into one folder header per project, named projects before No project", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [
        linkRow({ key: "a1", project: "Amasty" }),
        linkRow({ key: "b1", project: "Balcia" }),
        linkRow({ key: "n1" }),
        linkRow({ key: "a2", project: "Amasty" }),
      ],
      () => undefined,
    );

    const trs = [...body.querySelectorAll("tr")];
    expect(trs.map(rowKind)).toEqual([
      "folder",
      "row",
      "row",
      "folder",
      "row",
      "folder",
      "row",
    ]);
    expect(
      [...body.querySelectorAll(".link-folder td")].map(
        (cell) => cell.textContent,
      ),
    ).toEqual(["Amasty · 2", "Balcia · 1", "No project · 1"]);
    expect(body.querySelectorAll("tr[data-key]")).toHaveLength(4);
  });

  it("gives a folder header row no checkbox and no data-key", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [linkRow({ key: "a1", project: "Amasty" })],
      () => undefined,
    );

    const header = body.querySelector(".link-folder") as HTMLTableRowElement;
    expect(header.querySelector("input")).toBeNull();
    expect(header.dataset.key).toBeUndefined();
  });

  it("clears stale rows on re-render", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [
        {
          key: "a",
          slide: 1,
          label: "x",
          source: "y",
          kind: "table",
          status: "current",
          pushedAt: 0,
          selected: true,
        },
      ],
      () => undefined,
    );
    renderLinkRows(body, [], () => undefined);
    expect(body.querySelectorAll("tr")).toHaveLength(0);
  });
});

describe("renderInbox", () => {
  it("lists items with insert buttons", () => {
    const list = document.createElement("div");
    const onInsert = vi.fn();
    const item = {
      id: "a".repeat(32),
      token: "t",
      kind: "range" as const,
      label: "Model!B4:F12",
      src: {
        workbook: "Model_v4.xlsx",
        sheet: "Model",
        ref: "B4:F12",
        anchor: "PLSFIX_LINK_aaaaaaaa",
      },
      createdAt: new Date().toISOString(),
    };
    renderInbox(list, [item], onInsert);
    list.querySelector("button")!.click();
    expect(onInsert).toHaveBeenCalledWith(item);
    renderInbox(list, [], onInsert);
    expect(list.textContent).toContain("Nothing waiting");
  });
});

describe("renderCandidates", () => {
  it("lists each export by label, workbook, kind and age, in the order given", () => {
    const select = document.createElement("select");

    renderCandidates(select, [waiting("Model_v5.xlsx"), waiting("Old.xlsx")]);

    const options = [...select.options];
    expect(options.map((option) => option.value)).toEqual([
      "Model_v5.xlsx",
      "Old.xlsx",
    ]);
    expect(options[0]!.textContent).toBe(
      "Model!B4:F12 · Model_v5.xlsx · Picture (your choice) · just now",
    );
    expect(options[0]!.title).toBe(options[0]!.textContent);
  });

  it("clears stale options on re-render", () => {
    const select = document.createElement("select");
    renderCandidates(select, [waiting("Model_v5.xlsx")]);
    renderCandidates(select, []);
    expect(select.options).toHaveLength(0);
  });
});

describe("labels", () => {
  it("names statuses", () => {
    expect(statusLabel("wrongKey")).toBe("Wrong link key");
  });

  // The relay answers 404 for a link it forgot after 30 days without a push
  // AND for one the modeller removed in Excel: the old "Source missing" sent
  // people hunting for a source that was never lost, and "Expired" would be
  // wrong for the removed one.
  it("says a link the relay no longer holds is not there, and why that can be", () => {
    expect(statusLabel("missing")).toBe("Not on the relay");
    expect(statusDetail("missing")).toBe(
      "Not on the relay: expired after 30 days without a push, or removed in Excel. Push it again, or export it again if it was removed.",
    );
  });

  it("gives every other status the same words as its tooltip", () => {
    for (const status of [
      "current",
      "updateAvailable",
      "wrongKey",
      "notPasted",
    ] as const) {
      expect(statusDetail(status)).toBe(statusLabel(status));
    }
  });

  it("shows the short badge and keeps the whole sentence in the tooltip and the accessible name", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(
      body,
      [linkRow({ key: "s1/sh1", status: "missing" })],
      vi.fn(),
    );

    const badge = body.querySelector(".link-status .badge") as HTMLElement;
    expect(badge.textContent).toBe("Not on the relay");
    expect(badge.title).toBe(
      "Not on the relay: expired after 30 days without a push, or removed in Excel. Push it again, or export it again if it was removed.",
    );
    expect(badge.getAttribute("aria-label")).toBe(badge.title);
    expect(badge.textContent).not.toContain("Source missing");
  });

  // The deck answered: the word, the time its data was sent under it, and the
  // whole sentence in the tooltip. 14:02 on 2 October of the current year.
  it("says In this deck with the time the deck's data was sent", () => {
    const sent = new Date(new Date().getFullYear(), 9, 2, 14, 2).getTime();
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;

    renderLinkRows(
      body,
      [linkRow({ key: "s1/sh1", status: "inDeck", pushedAt: sent / 1000 })],
      vi.fn(),
    );

    const cell = body.querySelector(".link-status") as HTMLElement;
    const badge = cell.querySelector(".badge") as HTMLElement;
    expect(badge.textContent).toBe("In this deck");
    expect(badge.classList.contains("indeck")).toBe(true);
    expect(cell.querySelector(".link-meta")!.textContent).toBe(
      "data from 02.10 14:02",
    );
    expect(badge.title).toContain("In this deck, data from 02.10 14:02. ");
    expect(badge.title).toContain("Copy or push it again from Excel");
    expect(badge.getAttribute("aria-label")).toBe(badge.title);
  });

  it("leaves the time line out for a deck row whose record carries no usable time", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;

    renderLinkRows(
      body,
      [linkRow({ key: "s1/sh1", status: "inDeck", pushedAt: null })],
      vi.fn(),
    );

    expect(body.querySelector(".link-status .link-meta")).toBeNull();
    expect(statusDetail("inDeck", null).startsWith("In this deck. ")).toBe(
      true,
    );
  });

  it("gives no other status a time line", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;

    renderLinkRows(
      body,
      [linkRow({ key: "s1/sh1", status: "current", pushedAt: 100 })],
      vi.fn(),
    );

    expect(body.querySelector(".link-status .link-meta")).toBeNull();
  });
});

describe("renderSlideOptions", () => {
  it("lists This slide plus one option per slide, in order", () => {
    const select = document.createElement("select");

    renderSlideOptions(select, 3);

    const options = [...select.options];
    expect(options.map((option) => option.value)).toEqual(["", "1", "2", "3"]);
    expect(options.map((option) => option.textContent)).toEqual([
      "This slide",
      "Slide 1",
      "Slide 2",
      "Slide 3",
    ]);
    expect(select.value).toBe("");
  });

  it("keeps the picked slide across a re-render, and falls back once it is gone", () => {
    const select = document.createElement("select");
    renderSlideOptions(select, 3);
    select.value = "2";

    renderSlideOptions(select, 3);
    expect(select.value).toBe("2");

    renderSlideOptions(select, 1);
    expect(select.value).toBe("");
  });
});
