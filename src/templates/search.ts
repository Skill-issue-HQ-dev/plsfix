// What typing into the Templates search keeps: an entry stays when every word of
// the query starts a word of its name, its one-line description or its group.
// Pure, so the picker's filter is proven without a DOM. Invariant: the order of
// the list it is given is the order of the answer; nothing is ranked.

interface Searchable {
  name: string;
  description: string;
  group: string;
}

// Lower-cased words, any run of letters and digits: a slash, a hyphen or a
// bracket is a gap, so "npv/irr" and "NPV / IRR" are the same two words.
function wordsOf(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((word) => word !== "");
}

export function filterTemplates<T extends Searchable>(
  entries: readonly T[],
  query: string,
): T[] {
  const wanted = wordsOf(query);
  if (wanted.length === 0) return [...entries];
  return entries.filter((entry) => {
    const have = wordsOf(`${entry.name} ${entry.description} ${entry.group}`);
    return wanted.every((word) =>
      have.some((candidate) => candidate.startsWith(word)),
    );
  });
}
