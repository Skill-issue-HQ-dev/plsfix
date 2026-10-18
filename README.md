# pls,fix

Named after the email every analyst knows.

pls,fix is a free, open-source add-in for Excel and PowerPoint, made for people who build financial models. In Excel it
formats a model in a few keystrokes, flags hard-coded numbers among formulas, rounds so the totals still add
up, and drops in ready-made valuation, LBO and operating blocks. In PowerPoint it keeps tables, charts and
numbers linked to the workbook, so one click updates every linked slide. It runs on Windows, Mac and the web.

Open source under the [MIT licence](LICENSE): read it, fork it, build on it.

## [Install: click here](https://github.com/Skill-issue-HQ-dev/plsfix/releases/latest)

## [Want to improve it? Click here](https://github.com/Skill-issue-HQ-dev/plsfix/issues/new)

## Build and test

You need Node.js 22 and a stable Rust toolchain.

```sh
npm ci
npm run build
npm run check
cargo test --manifest-path server/Cargo.toml
```

`npm run check` runs every gate: types, lint, format, the TypeScript tests, the Rust tests of the server and the demo generator, and the manifests. The last line runs the server's tests alone. [CONTRIBUTING.md](CONTRIBUTING.md) maps the source.
