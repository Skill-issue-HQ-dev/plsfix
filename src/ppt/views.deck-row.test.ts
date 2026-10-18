// @vitest-environment jsdom
// The "Not on a slide" row of the Linked objects table: a record the deck holds
// for an object that is no longer on any slide, with the two buttons that bring
// it back. Drawn into the same folders as the shapes' rows.

import { describe, expect, it, vi } from "vitest";
import { OFF_SLIDE_WORD, type DeckRowView } from "./deck-rows";
import { renderLinkRows, type LinkRowView } from "./views";

const ID = "0123456789abcdef0123456789abcdef";

function deckRow(overrides: Partial<DeckRowView> = {}): DeckRowView {
  return {
    offSlide: true,
    key: `deck/${ID}`,
    id: ID,
    label: "Model!B4:F12 table",
    source: "Model_v4.xlsx",
    kind: "table",
    pushedAt: new Date(Date.now() - 3 * 24 * 3600 * 1000).toISOString(),
    canPutBack: true,
    ...overrides,
  };
}

function shapeRow(
  overrides: Partial<LinkRowView> & { key: string },
): LinkRowView {
  return {
    slide: 2,
    label: "Model!A1",
    source: "Model.xlsx",
    kind: "range",
    status: "current",
    pushedAt: null,
    selected: false,
    ...overrides,
  };
}

function draw(
  rows: LinkRowView[],
  deck: DeckRowView[],
  handlers = { onPutBack: vi.fn(), onRemove: vi.fn() },
): HTMLTableSectionElement {
  document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
  const body = document.getElementById("rows") as HTMLTableSectionElement;
  renderLinkRows(body, rows, () => undefined, { rows: deck, ...handlers });
  return body;
}

describe("a row for a record no shape carries", () => {
  it("reads Not on a slide, names the link and its source, and has no tick and no slide", () => {
    const body = draw([], [deckRow()]);
    const row = body.querySelector("tr.link-offslide") as HTMLElement;
    expect(row.dataset.deckId).toBe(ID);
    // Not a shape row: nothing to tick, nothing the toolbar acts on.
    expect(row.hasAttribute("data-key")).toBe(false);
    expect(row.querySelector("input")).toBeNull();
    expect(row.querySelector(".link-slide")!.textContent).toBe("-");
    expect(row.querySelector(".link-label")!.firstChild!.textContent).toBe(
      "Model!B4:F12 table",
    );
    expect(row.querySelector(".link-meta")!.textContent).toBe(
      "Model_v4.xlsx · Table",
    );
    const badge = row.querySelector(".link-status .badge") as HTMLElement;
    expect(badge.textContent).toBe(OFF_SLIDE_WORD);
    expect(badge.title).toMatch(/still holds its link and data/);
    expect(badge.getAttribute("aria-label")).toBe(badge.title);
    expect(row.querySelector(".link-updated")!.textContent).toBe("3 days ago");
  });

  it("has exactly two buttons, each named for its link and keyed by its action", () => {
    const row = draw([], [deckRow()]).querySelector("tr.link-offslide")!;
    const buttons = [...row.querySelectorAll<HTMLButtonElement>("button")];
    expect(buttons.map((one) => one.dataset.action)).toEqual([
      "deck-put-back",
      "deck-remove",
    ]);
    expect(buttons.map((one) => one.textContent)).toEqual([
      "Put back",
      "Remove from this deck",
    ]);
    expect(buttons.map((one) => one.getAttribute("aria-label"))).toEqual([
      "Put back Model!B4:F12 table",
      "Remove Model!B4:F12 table from this deck",
    ]);
    expect(buttons.every((one) => one.type === "button")).toBe(true);
    expect(buttons.every((one) => !one.disabled)).toBe(true);
  });

  it("calls the pane with the link's id when Put back is pressed", () => {
    const handlers = { onPutBack: vi.fn(), onRemove: vi.fn() };
    const body = draw([], [deckRow()], handlers);
    body
      .querySelector<HTMLButtonElement>('[data-action="deck-put-back"]')!
      .click();
    expect(handlers.onPutBack).toHaveBeenCalledExactlyOnceWith(ID);
    expect(handlers.onRemove).not.toHaveBeenCalled();
  });

  it("removes only on the second press: the first one arms the button", () => {
    const handlers = { onPutBack: vi.fn(), onRemove: vi.fn() };
    const body = draw([], [deckRow()], handlers);
    const remove = body.querySelector<HTMLButtonElement>(
      '[data-action="deck-remove"]',
    )!;

    remove.click();
    expect(handlers.onRemove).not.toHaveBeenCalled();
    expect(remove.classList.contains("armed")).toBe(true);
    expect(remove.textContent).toBe("Click again to confirm");

    remove.click();
    expect(handlers.onRemove).toHaveBeenCalledExactlyOnceWith(ID);
    expect(remove.textContent).toBe("Remove from this deck");
    expect(handlers.onPutBack).not.toHaveBeenCalled();
  });

  it("says the deck kept no data, and cannot be put back, when only the identity was kept", () => {
    const why = "This deck kept no data for it. Send it again from Excel.";
    const body = draw(
      [],
      [
        deckRow({
          canPutBack: false,
          kind: "range",
          deadWhy: why,
          deadTag: "Data not kept in this deck",
        }),
      ],
    );
    const row = body.querySelector("tr.link-offslide")!;
    // The source line is as for any row; the reason has a line of its own.
    const [source, reason] = [...row.querySelectorAll(".link-meta")];
    expect(source!.textContent).toBe("Model_v4.xlsx · Picture (your choice)");
    expect(reason!.textContent).toBe("Data not kept in this deck");
    expect(reason!.getAttribute("title")).toBe(why);
    const [put, remove] = [
      ...row.querySelectorAll<HTMLButtonElement>("button"),
    ];
    expect(put!.disabled).toBe(true);
    expect(put!.title).toBe(why);
    // The record can still go: its link key is what the removal is for.
    expect(remove!.disabled).toBe(false);
    expect(remove!.dataset.dead).toBeUndefined();
    expect((row.querySelector(".badge") as HTMLElement).title).toMatch(
      /kept no data/,
    );
  });

  it("says why a chart this host cannot draw is dead, in the kind's own sentence", () => {
    const why =
      "Scatter charts are not drawn as shapes yet. Send it again from Excel: the Inbox inserts it as a picture.";
    const body = draw(
      [],
      [
        deckRow({
          canPutBack: false,
          kind: "chart",
          deadWhy: why,
          deadTag: "Cannot be drawn here",
        }),
      ],
    );
    const row = body.querySelector("tr.link-offslide")!;
    const [source, reason] = [...row.querySelectorAll(".link-meta")];
    expect(source!.textContent).toBe("Model_v4.xlsx · Chart");
    expect(reason!.textContent).toBe("Cannot be drawn here");
    const put = row.querySelector<HTMLButtonElement>(
      '[data-action="deck-put-back"]',
    )!;
    expect(put.disabled).toBe(true);
    expect(put.title).toBe(why);
    // The pane's busy latch keeps a dead button dead through this mark.
    expect(put.dataset.dead).toBe("true");
    const badge = row.querySelector(".badge") as HTMLElement;
    expect(badge.title).toContain(why);
    expect(`${badge.title} ${put.title}`).not.toMatch(/this PowerPoint/);
  });

  it("draws nothing for a record when the pane passes no rows of that kind", () => {
    document.body.innerHTML = "<table><tbody id='rows'></tbody></table>";
    const body = document.getElementById("rows") as HTMLTableSectionElement;
    renderLinkRows(body, [shapeRow({ key: "s/1" })], () => undefined);
    expect(body.querySelector(".link-offslide")).toBeNull();
    expect(body.querySelectorAll("tr[data-key]")).toHaveLength(1);
  });
});

describe("the rows among the shapes' rows", () => {
  it("joins its project's folder after that folder's shape rows, and counts in it", () => {
    const body = draw(
      [
        shapeRow({ key: "a1", project: "Amasty" }),
        shapeRow({ key: "n1" }),
        shapeRow({ key: "a2", project: "Amasty" }),
      ],
      [deckRow({ project: "Amasty" }), deckRow({ id: "f".repeat(32) })],
    );
    const lines = [...body.querySelectorAll("tr")].map((tr) =>
      tr.classList.contains("link-folder")
        ? `folder ${tr.textContent}`
        : (tr.getAttribute("data-key") ??
          `deck ${tr.getAttribute("data-deck-id")}`),
    );
    expect(lines).toEqual([
      "folder Amasty · 3",
      "a1",
      "a2",
      `deck ${ID}`,
      "folder No project · 2",
      "n1",
      `deck ${"f".repeat(32)}`,
    ]);
  });

  it("leaves shape rows exactly as they were drawn without any", () => {
    const rows = [shapeRow({ key: "s/1" }), shapeRow({ key: "s/2" })];
    document.body.innerHTML =
      "<table><tbody id='plain'></tbody><tbody id='both'></tbody></table>";
    const plain = document.getElementById("plain") as HTMLTableSectionElement;
    const both = document.getElementById("both") as HTMLTableSectionElement;
    renderLinkRows(plain, rows, () => undefined);
    renderLinkRows(both, rows, () => undefined, {
      rows: [],
      onPutBack: () => undefined,
      onRemove: () => undefined,
    });
    expect(both.innerHTML).toBe(plain.innerHTML);
  });

  it("clears the previous draw, so a stale button never survives", () => {
    const handlers = { onPutBack: vi.fn(), onRemove: vi.fn() };
    const body = draw([], [deckRow()], handlers);
    renderLinkRows(body, [], () => undefined, {
      rows: [],
      ...handlers,
    });
    expect(body.querySelector("button")).toBeNull();
  });
});
