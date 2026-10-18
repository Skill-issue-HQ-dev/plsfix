// The syntax half of the test formula evaluator: Excel formula text to a tree
// (numbers, strings, references, ranges, unary minus, the operator levels, calls).
// Test support only; test/formula-eval.ts evaluates the tree. Invariant: a token
// or shape it cannot read throws by name instead of being guessed at.

type Token =
  | { kind: "num"; value: number }
  | { kind: "str"; value: string }
  | { kind: "name"; value: string }
  | { kind: "op"; value: string };

export type Node =
  | { type: "num"; value: number }
  | { type: "str"; value: string }
  | { type: "ref"; address: string }
  | { type: "range"; from: string; to: string }
  | { type: "neg"; operand: Node }
  | { type: "bin"; op: string; left: Node; right: Node }
  | { type: "call"; name: string; args: Node[] };

const NUMBER = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?/i;
const NAME = /^[A-Za-z_$][A-Za-z0-9_.$]*/;
const OPERATOR = /^(<=|>=|<>|[-+*/^(),:<>=])/;

function readString(source: string, start: number): [string, number] {
  let text = "";
  let at = start + 1;
  while (at < source.length) {
    if (source[at] === '"' && source[at + 1] === '"') {
      text += '"';
      at += 2;
    } else if (source[at] === '"') {
      return [text, at + 1];
    } else {
      text += source[at];
      at += 1;
    }
  }
  throw new Error(`unterminated string in ${source}`);
}

function tokenize(source: string): Token[] {
  const tokens: Token[] = [];
  let at = 0;
  while (at < source.length) {
    const rest = source.slice(at);
    const number = NUMBER.exec(rest);
    const name = NAME.exec(rest);
    const operator = OPERATOR.exec(rest);
    if (/^\s/.test(rest)) {
      at += 1;
    } else if (rest.startsWith('"')) {
      const [value, next] = readString(source, at);
      tokens.push({ kind: "str", value });
      at = next;
    } else if (number) {
      tokens.push({ kind: "num", value: Number(number[0]) });
      at += number[0].length;
    } else if (name) {
      tokens.push({ kind: "name", value: name[0] });
      at += name[0].length;
    } else if (operator) {
      tokens.push({ kind: "op", value: operator[0] });
      at += operator[0].length;
    } else {
      throw new Error(`cannot read "${rest}" in ${source}`);
    }
  }
  return tokens;
}

class Parser {
  private at = 0;
  constructor(private readonly tokens: Token[]) {}

  parse(): Node {
    const node = this.comparison();
    if (this.at < this.tokens.length) {
      throw new Error(
        `unexpected token ${JSON.stringify(this.tokens[this.at])}`,
      );
    }
    return node;
  }

  private peekOp(...ops: string[]): string | null {
    const token = this.tokens[this.at];
    return token?.kind === "op" && ops.includes(token.value)
      ? token.value
      : null;
  }

  private binary(next: () => Node, ...ops: string[]): Node {
    let left = next();
    for (let op = this.peekOp(...ops); op !== null; op = this.peekOp(...ops)) {
      this.at += 1;
      left = { type: "bin", op, left, right: next() };
    }
    return left;
  }

  private comparison(): Node {
    return this.binary(() => this.additive(), "=", "<>", "<=", ">=", "<", ">");
  }

  private additive(): Node {
    return this.binary(() => this.term(), "+", "-");
  }

  private term(): Node {
    return this.binary(() => this.power(), "*", "/");
  }

  // Excel's ^ is left-associative and binds looser than a unary minus.
  private power(): Node {
    return this.binary(() => this.unary(), "^");
  }

  private unary(): Node {
    if (this.peekOp("-")) {
      this.at += 1;
      return { type: "neg", operand: this.unary() };
    }
    if (this.peekOp("+")) {
      this.at += 1;
      return this.unary();
    }
    return this.primary();
  }

  private primary(): Node {
    const token = this.tokens[this.at];
    this.at += 1;
    if (token?.kind === "num") return { type: "num", value: token.value };
    if (token?.kind === "str") return { type: "str", value: token.value };
    if (token?.kind === "op" && token.value === "(") {
      const inner = this.comparison();
      this.expectOp(")");
      return inner;
    }
    if (token?.kind === "name") return this.nameNode(token.value);
    throw new Error(`unexpected token ${JSON.stringify(token)}`);
  }

  private nameNode(name: string): Node {
    if (this.peekOp("(")) {
      this.at += 1;
      const args: Node[] = [];
      while (!this.peekOp(")")) {
        args.push(this.comparison());
        if (this.peekOp(",")) this.at += 1;
      }
      this.expectOp(")");
      return { type: "call", name: name.toUpperCase(), args };
    }
    if (this.peekOp(":")) {
      this.at += 1;
      const end = this.tokens[this.at];
      this.at += 1;
      if (end?.kind !== "name") throw new Error(`range of ${name} has no end`);
      return { type: "range", from: name, to: end.value };
    }
    return { type: "ref", address: name };
  }

  private expectOp(op: string): void {
    if (!this.peekOp(op)) throw new Error(`expected ${op}`);
    this.at += 1;
  }
}

/** A formula's text, without its leading "=", as a syntax tree. */
export function parseFormula(source: string): Node {
  return new Parser(tokenize(source)).parse();
}
