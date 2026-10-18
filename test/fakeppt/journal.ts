// The undo log of what a batch has written since the last sync, off the model:
// dropped when a sync lands (committed copies catch up), played backwards when
// one fails or never answers, played back from a mark when the batch halts at
// an op the host cannot do. Invariant: an undo restores exactly what its own
// write changed, nothing else.

const logs = new WeakMap<object, (() => void)[]>();
const commits = new WeakMap<object, (() => void)[]>();

// Called by a write, with the closure that takes it back.
export function journalRecord(deck: object, undo: () => void): void {
  const log = logs.get(deck);
  if (log) log.push(undo);
  else logs.set(deck, [undo]);
}

// A store that keeps what the host holds apart from what a batch has queued
// registers here once, and is called after every landed sync that wrote.
export function journalOnConfirm(deck: object, commit: () => void): void {
  const list = commits.get(deck);
  if (list) list.push(commit);
  else commits.set(deck, [commit]);
}

// A sync that landed, or was refused after the host applied the batch.
export function confirmJournal(deck: object): void {
  if (!logs.delete(deck)) return;
  for (const commit of commits.get(deck) ?? []) commit();
}

// A sync that failed or never answered: last write undone first.
export function rollbackJournal(deck: object): void {
  const log = logs.get(deck) ?? [];
  logs.delete(deck);
  for (const undo of log.reverse()) undo();
}

// What a batch halted at a failing op must be able to take back although a sync
// that fails never does (the fake applies those writes when they are queued): a
// shape's delete, a tag's add or delete. Kept apart, so the sync that fails or
// never answers rolls back exactly what it always did.
const halts = new WeakMap<object, (() => void)[]>();

export function haltRecord(deck: object, undo: () => void): void {
  const log = halts.get(deck);
  if (log) log.push(undo);
  else halts.set(deck, [undo]);
}

export function journalMark(deck: object): number {
  return logs.get(deck)?.length ?? 0;
}

export function haltMark(deck: object): number {
  return halts.get(deck)?.length ?? 0;
}

// The writes after `mark`, played backwards: the part of a batch that never ran.
export function rollbackJournalAfter(deck: object, mark: number): void {
  for (const undo of (logs.get(deck) ?? []).splice(mark).reverse()) undo();
}

export function haltRollbackAfter(deck: object, mark: number): void {
  for (const undo of (halts.get(deck) ?? []).splice(mark).reverse()) undo();
}

// A sync ended the batch, whichever way: nothing is left to halt.
export function haltClear(deck: object): void {
  halts.delete(deck);
}
