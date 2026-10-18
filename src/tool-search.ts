// Ranks the pane's tools by what they do: typing part of a label, part of its
// help sentence or a word of its section's heading should find the button, not
// just an exact name match. Pure so it is unit-tested without a DOM; the
// catalogue that turns the real markup into ToolEntry[] lives in
// src/pane/tool-search.ts.

/** One searchable tool: a button's identity, where it lives, what it does. */
export interface ToolEntry {
  /** The button's data-action, else its id - what going to the tool looks up. */
  action: string;
  label: string;
  tab: string;
  sentence: string;
  /** The h2 of the section the button sits in, and the kicker above it. */
  heading?: string;
  kicker?: string;
}

const NAME_WEIGHT = 3;
const ALIAS_WEIGHT = 2;
const SENTENCE_WEIGHT = 1;
const WHOLE_LABEL_PREFIX_BONUS = 5;

// The Templates picker's Insert buttons carry `template-` + the block's id as
// their data-action. They write a block into the sheet, so on equal scores
// every other row comes first.
const TEMPLATE_ACTION_PREFIX = "template-";

// Words a query carries that no label or sentence is looking for.
const STOP_WORDS: ReadonlySet<string> = new Set([
  "a",
  "an",
  "the",
  "to",
  "of",
  "for",
  "my",
  "in",
  "on",
  "and",
  "with",
]);

// What a query word also means, as words of a label: a modeller types what they
// want done, the button says what it is called. Data, not logic: a pair added
// here is all a new synonym takes. A label word only, never a sentence. A word
// that is already its own label ("undo") is listed too, so every word the pane
// answers by name is in one place; its own pair changes nothing.
const ALIASES: ReadonlyMap<string, readonly string[]> = new Map([
  ["trace", ["precedents", "dependents"]],
  ["send", ["export", "copy", "push"]],
  ["update", ["export", "copy", "push"]],
  ["powerpoint", ["export", "copy", "push"]],
  ["undo", ["undo"]],
  ["multiple", ["0.0x"]],
]);

// The pane spells US English, so a query typed the British way is read as the
// word its cards use; the pair is applied to the lower-cased query, plurals too.
const SPELLINGS: readonly (readonly [string, string])[] = [["colour", "color"]];

function queryOf(query: string): string {
  return SPELLINGS.reduce(
    (text, [british, us]) => text.replaceAll(british, us),
    query.toLowerCase(),
  );
}

function wordsOf(text: string): string[] {
  return text.split(/\s+/).filter((word) => word.length > 0);
}

// The words that count: the stop words go, unless that would leave nothing.
// A query of "to" is a word typed halfway ("tornado"), not a question.
function tokensOf(query: string): string[] {
  const words = wordsOf(queryOf(query));
  const meaningful = words.filter((word) => !STOP_WORDS.has(word));
  return meaningful.length > 0 ? meaningful : words;
}

interface Row {
  tool: ToolEntry;
  label: string;
  sentence: string;
  section: string;
}

interface Hit {
  weight: number;
  /** True when the label itself carries the token or an alias of it. */
  label: boolean;
}

function startsAWord(text: string, token: string): boolean {
  return wordsOf(text).some((word) => word.startsWith(token));
}

// The best way a token reaches a row. A word of the label, or of the section's
// heading or kicker, is a name hit; a label word the token stands for (the alias
// table) is worth a little less; the help sentence least. A name hit also takes
// a plain substring of the label; a section only ever matches the start of one
// of its words, because one heading word would otherwise pull in every row of
// its section.
function hitOf(token: string, row: Row): Hit | null {
  const direct = startsAWord(row.label, token) || row.label.includes(token);
  const alias =
    !direct &&
    (ALIASES.get(token) ?? []).some((word) => startsAWord(row.label, word));
  const inSection = startsAWord(row.section, token);
  const inSentence =
    startsAWord(row.sentence, token) || row.sentence.includes(token);
  if (direct || inSection) {
    return { weight: NAME_WEIGHT, label: direct || alias };
  }
  if (alias) return { weight: ALIAS_WEIGHT, label: true };
  if (inSentence) return { weight: SENTENCE_WEIGHT, label: false };
  return null;
}

interface Ranked {
  tool: ToolEntry;
  score: number;
  label: boolean;
  template: boolean;
}

function rankedOf(tokens: string[], row: Row): Ranked | null {
  let score = 0;
  let label = false;
  for (const token of tokens) {
    const hit = hitOf(token, row);
    if (hit === null) return null;
    score += hit.weight;
    label ||= hit.label;
  }
  if (row.label.startsWith(tokens.join(" "))) {
    score += WHOLE_LABEL_PREFIX_BONUS;
    label = true;
  }
  return {
    tool: row.tool,
    score,
    label,
    template: row.tool.action.startsWith(TEMPLATE_ACTION_PREFIX),
  };
}

/**
 * Every tool whose label, sentence or section carries all of the query's words
 * (an alias counts for a label word), best match first. A tool missing even one
 * word is left out entirely - this is an AND filter, not a fuzzy score over
 * however many words happened to hit. On equal scores a template Insert row
 * comes after every other row, then a row whose label hits comes before one that
 * only its sentence or section does, then the order `tools` was given in
 * (Array#sort is stable). An empty or whitespace-only query matches nothing.
 */
export function rankTools(
  query: string,
  tools: ToolEntry[],
  limit = 8,
): ToolEntry[] {
  const tokens = tokensOf(query);
  if (tokens.length === 0) return [];

  const ranked: Ranked[] = [];
  for (const tool of tools) {
    const row: Row = {
      tool,
      label: tool.label.toLowerCase(),
      sentence: tool.sentence.toLowerCase(),
      section: `${tool.heading ?? ""} ${tool.kicker ?? ""}`.toLowerCase(),
    };
    const entry = rankedOf(tokens, row);
    if (entry !== null) ranked.push(entry);
  }

  ranked.sort(
    (a, b) =>
      b.score - a.score ||
      Number(a.template) - Number(b.template) ||
      Number(b.label) - Number(a.label),
  );
  return ranked.slice(0, limit).map((entry) => entry.tool);
}

/**
 * `text` cut at the last whole word that fits in `max` characters, the
 * ellipsis included, never in the middle of a word and never leaving a comma
 * or a colon hanging. A word longer than the room is cut hard: there is no
 * boundary to use.
 */
export function clipAtWord(text: string, max: number): string {
  if (text.length <= max) return text;
  const room = max - 1;
  const space = text.slice(0, room + 1).lastIndexOf(" ");
  const cut = space > 0 ? space : room;
  return `${text.slice(0, cut).replace(/[\s,;:.(-]+$/, "")}…`;
}

/** The one line a query with no hit gets, in the tone of the Templates box's own. */
export function noToolLine(query: string): string {
  return `No tool matches "${query.trim()}". Try export, undo, color or dcf.`;
}
