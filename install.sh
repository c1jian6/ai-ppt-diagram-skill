#!/usr/bin/env bash
# Install this skill into the agent's skill directory (one line):
#
#   curl -fsSL https://raw.githubusercontent.com/C1jian6/ai-ppt-diagram-skill/main/install.sh | bash
#
# Override the target with:  bash install.sh ~/some/skills/dot2pptx
# Override the source with:  DOT2PPTX_REPO=https://github.com/you/fork.git bash install.sh

set -euo pipefail

REPO="${DOT2PPTX_REPO:-https://github.com/C1jian6/ai-ppt-diagram-skill.git}"
DEST="${1:-$HOME/.dsh/skills/dot2pptx}"

command -v git >/dev/null 2>&1 || { echo "git is required"; exit 1; }

if [ -d "$DEST/.git" ]; then
  echo "updating $DEST"
  git -C "$DEST" pull --ff-only
elif [ -d "$DEST" ]; then
  echo "$DEST exists but is not a git checkout; remove it first"; exit 1
else
  echo "cloning into $DEST"
  mkdir -p "$(dirname "$DEST")"
  git clone --depth 1 "$REPO" "$DEST"
fi

echo
echo "checking the environment:"
python3 "$DEST/scripts/bootstrap.py" --check || true
echo
echo "done. restart the agent session so the skill is picked up."
