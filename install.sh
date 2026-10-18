#!/bin/sh
# Copy office-read and office-edit into a skills directory (overwrites previous copies).
# Usage: ./install.sh [destination]   default: ~/.claude/skills
set -e
src="$(cd "$(dirname "$0")" && pwd)"
dst="${1:-$HOME/.claude/skills}"
mkdir -p "$dst"
for s in office-read office-edit; do
  rm -rf "$dst/$s"
  cp -r "$src/$s" "$dst/$s"
  find "$dst/$s" -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
  echo "installed $s -> $dst/$s"
done
echo "Start a new agent session to pick up the skills."
