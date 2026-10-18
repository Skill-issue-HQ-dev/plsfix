// The rules of the picture-era upgrade, one row each: what a picture link may
// become when its source changed, which ones never do, how a batch isolates the
// upgrades, and why a chart link that stays a picture says so. The host comes
// from the fake deck, so "this host cannot draw it" is a real API answer.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { columnChart, bootPpt } from "../../test/ppt.support";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
} from "../../test/fakeppt";
import type { ChartData } from "../link/chart-model";
import type {
  LinkKind,
  Payload,
  PicturePayload,
  TablePayload,
} from "../link/model";
import { TABLE_MAX_ROWS } from "../link/model";
import type { FoundLink } from "./host";
import type { RefreshRequest } from "./refresh";
import {
  isChosen,
  isolateUpgrades,
  upgradeOf,
  UNGROUP_FOR_CHART,
  wantsUpgrade,
  whyNotNative,
} from "./upgrade";

enableStrictLoadSemantics();

let helpers: FakePptHelpers;
beforeEach(async () => {
  ({ helpers } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

const SRC = { workbook: "M.xlsx", sheet: "Model", ref: "B4:F12", anchor: "a" };

function found(kind: LinkKind, over: Partial<FoundLink> = {}): FoundLink {
  return {
    slideId: "slide-1",
    slideIndex: 0,
    shapeId: "shape-1",
    tag: { v: 1, id: "0".repeat(32), kind, rev: 1, src: SRC, pushedAt: "x" },
    token: "t",
    type: "GeometricShape",
    left: 10,
    top: 10,
    width: 400,
    height: 200,
    ...over,
  };
}

function chartPayload(chart?: ChartData, issue?: string): PicturePayload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png: fakePng(800, 400),
    src: SRC,
    pushedAt: "y",
    hash: "0".repeat(64),
    ...(chart === undefined ? {} : { chart }),
    ...(issue === undefined ? {} : { chartIssue: issue }),
  };
}

function tablePayload(rows = 2, cols = 2): TablePayload {
  const cells = Array.from({ length: rows }, () =>
    Array.from({ length: cols }, () => ({ t: "1" })),
  );
  return {
    v: 1,
    kind: "table",
    rows,
    cols,
    cells,
    widths: Array.from({ length: cols }, () => 50),
    src: SRC,
    pushedAt: "y",
    hash: "0".repeat(64),
  };
}

const COLUMN = columnChart(6);
const PIE: ChartData = { ...columnChart(3), kind: "pie" };

describe("upgradeOf: a chart link that is a picture", () => {
  it("becomes a chart when the payload carries drawable data", () => {
    const up = upgradeOf(found("chart"), chartPayload(COLUMN));
    expect(up?.to).toBe("chart");
  });

  it.each([
    ["a picture the user chose", found("chart", { tag: chosenTag() })],
    ["a range link's picture", found("range")],
    ["a native chart already", found("chart", { type: "Group" })],
    [
      "a picture inside a group of the user's",
      found("chart", { groupPath: ["g"] }),
    ],
  ])("stays a picture: %s", (_name, link) => {
    expect(upgradeOf(link, chartPayload(COLUMN))).toBeNull();
  });

  it("stays a picture when the payload carries no chart data", () => {
    expect(
      upgradeOf(found("chart"), chartPayload(undefined, "no data")),
    ).toBeNull();
  });

  it("stays a picture on a host that cannot draw the kind", () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    expect(upgradeOf(found("chart"), chartPayload(PIE))).toBeNull();
    helpers.setSupported((_set, version) => version === "1.7");
    expect(upgradeOf(found("chart"), chartPayload(COLUMN))).toBeNull();
  });

  function chosenTag(): FoundLink["tag"] {
    return { ...found("chart").tag, choice: "picture" };
  }

  it("reads the choice off the tag", () => {
    expect(isChosen(found("chart", { tag: chosenTag() }))).toBe(true);
    expect(isChosen(found("chart"))).toBe(false);
  });
});

describe("upgradeOf: a range link that is a picture", () => {
  it("becomes a table only when the payload is one", () => {
    expect(upgradeOf(found("range"), tablePayload())?.to).toBe("table");
    expect(upgradeOf(found("range"), chartPayload())).toBeNull();
  });

  it("stays a picture past the table cap or without the table API", () => {
    expect(
      upgradeOf(found("range"), tablePayload(TABLE_MAX_ROWS + 1)),
    ).toBeNull();
    helpers.setSupported((_set, version) => version !== "1.8");
    expect(upgradeOf(found("range"), tablePayload())).toBeNull();
  });

  it("never touches a table, a text or a native table link", () => {
    expect(
      upgradeOf(found("table", { type: "Table" }), tablePayload()),
    ).toBeNull();
    expect(
      upgradeOf(found("text", { type: "TextBox" }), tablePayload()),
    ).toBeNull();
  });
});

describe("isolateUpgrades", () => {
  const request = (link: FoundLink, payload: Payload, upgrade?: boolean) =>
    ({ found: link, payload, rev: 2, upgrade }) as RefreshRequest;

  it("sets each upgrade on its own and keeps the rest together", () => {
    const plain = request(found("range"), chartPayload(), true);
    const other = request(
      found("range", { shapeId: "s2" }),
      chartPayload(),
      true,
    );
    const up = request(
      found("chart", { shapeId: "s3" }),
      chartPayload(COLUMN),
      true,
    );
    expect(wantsUpgrade(up)).toBe(true);
    expect(wantsUpgrade(plain)).toBe(false);
    expect(isolateUpgrades([[plain, up, other]])).toEqual([
      [plain, other],
      [up],
    ]);
    expect(isolateUpgrades([[plain, other]])).toEqual([[plain, other]]);
  });

  it("leaves a request that did not ask for an upgrade (a revert) in the batch", () => {
    const revert = request(found("chart"), chartPayload(COLUMN));
    expect(wantsUpgrade(revert)).toBe(false);
    expect(isolateUpgrades([[revert]])).toEqual([[revert]]);
  });
});

describe("whyNotNative", () => {
  it("is silent for a picture the user chose and for a range picture", () => {
    const chosen = { ...found("chart").tag, choice: "picture" as const };
    expect(
      whyNotNative(found("chart", { tag: chosen }), chartPayload(PIE)),
    ).toBeUndefined();
    expect(whyNotNative(found("range"), chartPayload(PIE))).toBeUndefined();
  });

  it("gives Excel's reason, this host's reason, or the group", () => {
    expect(whyNotNative(found("chart"), chartPayload(undefined, "x"))).toMatch(
      /^as a picture: x/,
    );
    helpers.setSupported((_set, version) => version !== "1.10");
    expect(whyNotNative(found("chart"), chartPayload(PIE))).toMatch(
      /pie shapes need/,
    );
    helpers.setSupported(() => true);
    expect(
      whyNotNative(found("chart", { groupPath: ["g"] }), chartPayload(COLUMN)),
    ).toBe(UNGROUP_FOR_CHART);
    expect(whyNotNative(found("chart"), chartPayload(COLUMN))).toBeUndefined();
  });
});
