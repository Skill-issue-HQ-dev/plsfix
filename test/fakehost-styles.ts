// Named cell styles for the fake Excel host (test/fakehost.ts wires this in): styles.add / getItem,
// a style's font, fill, borders, number format, alignment and include flags, range.style, an edit
// reaching every wearer, delete sending them to Normal. A style's body lives in a side table.
// Simplification: an edit overwrites the included categories of every wearer (Excel keeps their own).

import {
  defaultCell,
  FillProxy,
  FontProxy,
  hostError,
  NORMAL_STYLE,
  type BorderEdge,
  type FakeBorder,
  type FakeCell,
  type FakeFill,
  type FakeFont,
  type FakeStyle,
  type FakeWorkbook,
} from "./fakehost";
import { copyCategories } from "./fakehost-categories";
import { checkRefusal } from "./fakehost-refusals";

export interface StyleInclude {
  number: boolean;
  font: boolean;
  patterns: boolean;
  border: boolean;
  alignment: boolean;
  protection: boolean;
}

export interface StyleBody {
  numberFormat: string;
  horizontalAlignment: string;
  font: FakeFont;
  fill: FakeFill;
  borders: Record<BorderEdge, FakeBorder>;
  locked: boolean;
  formulaHidden: boolean;
  include: StyleInclude;
}

/** What a proxy needs of the context: the way a refusal reaches the sync. */
interface Refusals {
  queueError(error: Error): void;
}

const INCLUDE_FLAGS = [
  "includeNumber",
  "includeFont",
  "includePatterns",
  "includeBorder",
  "includeAlignment",
  "includeProtection",
] as const;
type IncludeFlag = (typeof INCLUDE_FLAGS)[number];
const FLAG_KEY: Record<IncludeFlag, keyof StyleInclude> = {
  includeNumber: "number",
  includeFont: "font",
  includePatterns: "patterns",
  includeBorder: "border",
  includeAlignment: "alignment",
  includeProtection: "protection",
};

// The edges a style's border collection answers; Excel has no inside edges here.
const STYLE_EDGE: Record<string, BorderEdge> = {
  EdgeTop: "top",
  EdgeBottom: "bottom",
  EdgeLeft: "left",
  EdgeRight: "right",
  DiagonalDown: "diagonalDown",
  DiagonalUp: "diagonalUp",
};

const bodies = new WeakMap<FakeStyle, StyleBody>();

// A fresh style is Excel's Normal: the look an untouched cell has, every
// category included.
function normalBody(): StyleBody {
  const cell = defaultCell();
  return {
    numberFormat: cell.numberFormat,
    horizontalAlignment: cell.horizontalAlignment,
    font: cell.font,
    fill: cell.fill,
    borders: cell.borders,
    locked: cell.locked,
    formulaHidden: cell.formulaHidden,
    include: {
      number: true,
      font: true,
      patterns: true,
      border: true,
      alignment: true,
      protection: true,
    },
  };
}

function bodyOf(record: FakeStyle): StyleBody {
  let body = bodies.get(record);
  if (!body) {
    body = normalBody();
    bodies.set(record, body);
  }
  return body;
}

const sameName = (first: string, second: string): boolean =>
  first.toLowerCase() === second.toLowerCase();

function findStyle(
  workbook: FakeWorkbook,
  name: string,
): FakeStyle | undefined {
  return workbook.styles.find((record) => sameName(record.name, name));
}

function wearers(workbook: FakeWorkbook, name: string): FakeCell[] {
  return workbook.sheets.flatMap((sheet) =>
    [...sheet.cells.values()].filter((cell) => sameName(cell.style, name)),
  );
}

// An edit to a style is an edit to every cell wearing it.
function followEdit(workbook: FakeWorkbook, record: FakeStyle): void {
  const body = bodyOf(record);
  for (const cell of wearers(workbook, record.name)) copyCategories(cell, body);
}

function wearStyle(cell: FakeCell, record: FakeStyle): void {
  copyCategories(cell, bodyOf(record));
  cell.style = record.name;
}

function unknownStyle(name: string): Error {
  return hostError(
    "InvalidArgument",
    `The style ${name} does not exist in this workbook.`,
  );
}

/** What range.style answers: the one name every cell wears, else null. */
export function uniformStyle(names: string[][]): string | null {
  const flat = names.flat();
  const first = flat[0] ?? null;
  return flat.every((name) => name === first) ? first : null;
}

/**
 * range.style = name over the cells `each` visits (the range's own write, so a
 * protected sheet refuses it the same way). A style the workbook does not have
 * is refused on the sync, with nothing written.
 */
export function applyNamedStyle(
  workbook: FakeWorkbook,
  ctx: Refusals,
  name: string,
  each: (write: (cell: FakeCell) => void) => void,
): void {
  const record = findStyle(workbook, name);
  if (!record) {
    ctx.queueError(unknownStyle(name));
    return;
  }
  each((cell) => wearStyle(cell, record));
}

/** The `style` of one cell's setCellProperties entry: before its format. */
export function writeCellStyle(
  workbook: FakeWorkbook,
  ctx: Refusals,
  cell: FakeCell,
  name: string,
): void {
  const record = findStyle(workbook, name);
  if (!record) ctx.queueError(unknownStyle(name));
  else wearStyle(cell, record);
}

// ---------------------------------------------------------------------------
// The style object
// ---------------------------------------------------------------------------

class StyleBorderProxy {
  constructor(
    private border: FakeBorder,
    private touched: () => void,
  ) {}

  load(): this {
    return this;
  }

  get style(): string {
    return this.border.style;
  }
  set style(value: string) {
    this.border.style = value;
    this.touched();
  }

  get color(): string {
    return this.border.color;
  }
  set color(value: string) {
    this.border.color = value;
    this.touched();
  }

  get weight(): string {
    return this.border.weight;
  }
  set weight(value: string) {
    this.border.weight = value;
    this.touched();
  }
}

class StyleProxy {
  isNullObject = false;

  constructor(
    private workbook: FakeWorkbook,
    private ctx: Refusals,
    private record: FakeStyle,
  ) {}

  load(): this {
    return this;
  }

  private get body(): StyleBody {
    return bodyOf(this.record);
  }

  // Every write reaches the cells that wear the style at once.
  private touched = (): void => {
    checkRefusal(this.workbook);
    followEdit(this.workbook, this.record);
  };

  get name(): string {
    return this.record.name;
  }

  get builtIn(): boolean {
    return this.record.builtIn;
  }

  get numberFormat(): string {
    return this.body.numberFormat;
  }
  set numberFormat(value: string) {
    this.body.numberFormat = value;
    this.touched();
  }

  get horizontalAlignment(): string {
    return this.body.horizontalAlignment;
  }
  set horizontalAlignment(value: string) {
    this.body.horizontalAlignment = value;
    this.touched();
  }

  get font(): FontProxy {
    return new FontProxy(
      () => this.body.font,
      (apply) => {
        apply(this.body.font);
        this.touched();
      },
    );
  }

  get fill(): FillProxy {
    return new FillProxy(
      () => this.body.fill,
      (apply) => {
        apply(this.body.fill);
        this.touched();
      },
    );
  }

  get borders(): { getItem(index: string): StyleBorderProxy } {
    const body = this.body;
    const touched = this.touched;
    return {
      getItem(index: string): StyleBorderProxy {
        const edge = STYLE_EDGE[index];
        if (!edge) {
          throw hostError("InvalidArgument", `A style has no ${index} border.`);
        }
        return new StyleBorderProxy(body.borders[edge], touched);
      },
    };
  }

  // Excel refuses to delete one of its own styles; a scrubber that offered one
  // would fail at the host rather than quietly do nothing. The cells that wore
  // the style go back to Normal in the categories it included.
  delete(): void {
    if (this.record.builtIn) {
      throw hostError(
        "InvalidArgument",
        `${this.record.name} is a built-in style.`,
      );
    }
    const list = this.workbook.styles;
    const index = list.indexOf(this.record);
    if (index < 0) {
      this.ctx.queueError(unknownStyle(this.record.name));
      return;
    }
    const normal = findStyle(this.workbook, NORMAL_STYLE);
    const reset: StyleBody = {
      ...(normal ? bodyOf(normal) : normalBody()),
      include: this.body.include,
    };
    for (const cell of wearers(this.workbook, this.record.name)) {
      copyCategories(cell, reset);
      cell.style = NORMAL_STYLE;
    }
    list.splice(index, 1);
  }
}

// One include flag per property, the way office.js names them.
for (const flag of INCLUDE_FLAGS) {
  Object.defineProperty(StyleProxy.prototype, flag, {
    get(this: { body: StyleBody }): boolean {
      return this.body.include[FLAG_KEY[flag]];
    },
    set(this: { body: StyleBody; touched: () => void }, value: boolean) {
      this.body.include[FLAG_KEY[flag]] = Boolean(value);
      this.touched();
    },
  });
}

/** workbook.styles: the collection office.js hands out. */
export function styleCollection(workbook: FakeWorkbook, ctx: Refusals) {
  const list = workbook.styles;
  const wrap = (record: FakeStyle) => new StyleProxy(workbook, ctx, record);
  return {
    load: () => undefined,
    get items(): StyleProxy[] {
      return list.map(wrap);
    },
    // Excel keeps style names unique whatever their case, its own included.
    add(name: string): void {
      if (name.trim() === "") {
        throw hostError("InvalidArgument", "A style needs a name.");
      }
      if (findStyle(workbook, name)) {
        throw hostError("ItemAlreadyExists", `${name} already exists.`);
      }
      const record: FakeStyle = { name, builtIn: false };
      list.push(record);
      bodyOf(record);
    },
    getItem(name: string): StyleProxy {
      const record = findStyle(workbook, name);
      if (!record) throw hostError("ItemNotFound", `No style ${name}.`);
      return wrap(record);
    },
    getItemOrNullObject(name: string): StyleProxy {
      const record = findStyle(workbook, name);
      if (record) return wrap(record);
      const missing = wrap({ name, builtIn: false });
      missing.isNullObject = true;
      return missing;
    },
  };
}
