// A local paste is an update like any other: the bundle Excel copied is ingested
// and the rows it names are repainted, so a picture-era chart or a range picture
// whose link now carries a table is upgraded by the paste, at the picture's
// corner and width. A picture nobody converted is repainted as a picture.
// Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { LocalCollector } from "../src/link/local-collector";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import {
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkKind,
  type LinkTag,
  type Payload,
} from "../src/link/model";
import { pasteLinks } from "../src/ppt/paste-links";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import { columnChart } from "./ppt.support";

enableStrictLoadSemantics();

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};
const PNG = fakePng(800, 400);
const CORNER = { left: 100, top: 80, width: 400, height: 200 };

let presentation: FakePresentation;
beforeEach(() => {
  uninstallFakePpt();
  const host = installFakePpt({ slides: 3 });
  presentation = host.presentation;
  host.helpers.selectSlide(presentation.slides[0]!.id);
});
afterEach(() => uninstallFakePpt());

// The picture a deck holds from an earlier version, tagged at revision 1.
function plantPicture(kind: LinkKind): { id: string; token: string } {
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const shape = presentation.addShape(presentation.slides[1]!, {
    type: "GeometricShape",
    fillImage: PNG,
    ...CORNER,
  });
  const tag: LinkTag = {
    v: 1,
    id,
    kind,
    rev: 1,
    src: SRC,
    pushedAt: new Date().toISOString(),
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
  return { id, token };
}

// What Excel's copy leaves on the clipboard for one link at revision 2.
async function copied(
  link: { id: string; token: string },
  payload: Payload,
): Promise<LocalCollector> {
  const collector = new LocalCollector();
  const keys = await deriveLinkKeys(link.token);
  await collector.putLink(
    link.id,
    keys.auth,
    await seal(keys.enc, link.id, encodePayload(payload)),
    1,
  );
  return collector;
}

const now = (): string => new Date().toISOString();

describe("pasting into a deck that holds pictures", () => {
  it("draws a picture-era chart as shapes at the picture's corner and width", async () => {
    const link = plantPicture("chart");
    const chart: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 800,
      height: 400,
      png: PNG,
      src: SRC,
      pushedAt: now(),
      hash: "1".repeat(64),
      chart: columnChart(6),
    };

    const summary = await pasteLinks(
      (await copied(link, chart)).bundle(),
      await LocalStore.open(memoryPersistence()),
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const [drawn, ...rest] = presentation.slides[1]!.shapes;
    expect(rest).toEqual([]);
    expect(drawn!.type).toBe("Group");
    expect(drawn).toMatchObject({ left: 100, top: 80, width: 400 });
    expect(summary.notes).toEqual([
      `${SRC.sheet}: ${SRC.ref} now a native chart`,
    ]);
  });

  it("builds a table where a range picture sat once its link carries one", async () => {
    const link = plantPicture("range");
    const table: Payload = {
      v: 1,
      kind: "table",
      rows: 1,
      cols: 2,
      cells: [[{ t: "Revenue" }, { t: "1 200" }]],
      widths: [90, 60],
      src: SRC,
      pushedAt: now(),
      hash: "1".repeat(64),
    };

    const summary = await pasteLinks(
      (await copied(link, table)).bundle(),
      await LocalStore.open(memoryPersistence()),
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const [built, ...rest] = presentation.slides[1]!.shapes;
    expect(rest).toEqual([]);
    expect(built!.type).toBe("Table");
    expect(built).toMatchObject({ left: 100, top: 80, width: 400 });
  });

  it("repaints a range picture as a picture while its link still carries one", async () => {
    const link = plantPicture("range");
    const picture: Payload = {
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 400,
      height: 200,
      png: fakePng(400, 200),
      src: SRC,
      pushedAt: now(),
      hash: "1".repeat(64),
    };

    await pasteLinks(
      (await copied(link, picture)).bundle(),
      await LocalStore.open(memoryPersistence()),
    );

    const [kept] = presentation.slides[1]!.shapes;
    expect(kept!.type).toBe("GeometricShape");
    expect(kept!.fillImage).toBe(fakePng(400, 200));
  });
});
