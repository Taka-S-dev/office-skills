#!/bin/sh
# Copy office-read and office-edit into skills directories (overwrites previous copies).
# Usage: ./install.sh [destination]
#   default: ~/.claude/skills, and ~/.copilot/skills when GitHub Copilot's folder exists
set -e
src="$(cd "$(dirname "$0")" && pwd)"
if [ $# -ge 1 ]; then
  set -- "$1"
else
  set -- "$HOME/.claude/skills"
  [ -d "$HOME/.copilot" ] && set -- "$@" "$HOME/.copilot/skills"
fi
for dst in "$@"; do
  mkdir -p "$dst"
  for s in office-read office-edit; do
    rm -rf "$dst/$s"
    cp -r "$src/$s" "$dst/$s"
    find "$dst/$s" -name __pycache__ -type d -exec rm -rf {} + 2>/dev/null || true
    echo "installed $s -> $dst/$s"
  done
done
echo "Start a new agent session to pick up the skills."
