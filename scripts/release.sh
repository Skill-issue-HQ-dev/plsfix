#!/bin/sh
# One reviewed merge = one version: bump (scripts/bump-patch.sh), commit "vX.Y.Z"
# and write the annotated tag in a single step, so no version lands untagged again.
# Refuses a dirty tree: the bump commit carries the bump and nothing else.
# The tag's message is the release's plain headline when one is given: the
# gateway's Telegram note shows it as the title above its bullets, which are the
# subjects of the version's product commits (so those say what the user gets).
set -eu
cd "$(dirname "$0")/.."
LEVEL="${1:-patch}"
HEADLINE="${2:-}"
# An explicit X.Y.Z sets the developer line by hand (07.10.2026: back on v3 after the 1.0 launch).
case "$LEVEL" in patch|minor|major|[0-9]*.[0-9]*.[0-9]*) ;; *) echo "usage: sh scripts/release.sh [patch|minor|major|X.Y.Z] [\"plain headline\"]"; exit 2;; esac
if [ -n "$(git status --porcelain)" ]; then
  echo "release: commit or stash your changes first (the bump commit must be only the bump)"; exit 1
fi
sh scripts/bump-patch.sh "$LEVEL"
V="$(node -p "require('./package.json').version")"
git add package.json package-lock.json server/Cargo.toml server/Cargo.lock manifest.xml manifest.prod.xml manifest.public.xml public/shortcuts.html
git commit -q -m "v$V"
git tag -a "v$V" -m "${HEADLINE:-v$V}"
echo "released v$V ($(git rev-parse --short HEAD)), tagged${HEADLINE:+: $HEADLINE}"
