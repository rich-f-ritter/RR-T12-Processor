#!/usr/bin/env bash
# Cut a fresh account-wide skill zip from the current checkout of main.
#
#   bash dev/package_skill.sh
#
# Produces <repo>/rr-t12-processor-skill.zip with SKILL.md at the zip ROOT (the
# shape the claude.ai uploader expects). Upload it at:
#   claude.ai -> Customize -> Skills -> upload -> enable
#
# The zip is git-ignored; it's a disposable, point-in-time snapshot of whatever
# is committed on this branch. Run it after pulling the latest main to ship the
# newest skill account-wide.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SKILL_DIR="$ROOT/.claude/skills/rr-t12-processor"
OUT="$ROOT/rr-t12-processor-skill.zip"

[ -f "$SKILL_DIR/SKILL.md" ] || { echo "ERROR: SKILL.md not found at $SKILL_DIR" >&2; exit 1; }

rm -f "$OUT"
( cd "$SKILL_DIR" && zip -qr "$OUT" . -x '*__pycache__*' '*.pyc' )

echo "packaged: $OUT"
echo "  source: $SKILL_DIR"
echo "  branch: $(git -C "$ROOT" rev-parse --abbrev-ref HEAD 2>/dev/null || echo n/a)"
echo "  commit: $(git -C "$ROOT" rev-parse --short HEAD 2>/dev/null || echo n/a)"
echo "  files:"
unzip -l "$OUT" | sed 's/^/    /'
