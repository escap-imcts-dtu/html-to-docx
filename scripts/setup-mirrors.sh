#!/usr/bin/env bash
# Configure the two remotes used by this fork: GitHub (origin) and ADO (ado).
# Run this once after cloning.
#
# We mirror to two remotes for resilience: if one is accidentally deleted
# or becomes inaccessible, the other holds the full history.
#
# - origin: https://github.com/escap-imcts-dtu/html-to-docx.git           (consumed by ESCAP)
# - ado:    https://unescap.visualstudio.com/ESCAP-Document-Center/_git/html-to-docx (backup)
#
# We deliberately do NOT configure origin to push to both URLs at once —
# that would require ADO creds to be set up in a git credential helper.
# Use ./scripts/push-mirrors.sh to push to both with az CLI auth.

set -euo pipefail

GH_URL="https://github.com/escap-imcts-dtu/html-to-docx.git"
ADO_URL="https://unescap.visualstudio.com/ESCAP-Document-Center/_git/html-to-docx"

# Make origin point at GitHub only (single push URL).
git remote set-url origin "$GH_URL"

# Add or update the `ado` remote.
if git remote get-url ado > /dev/null 2>&1; then
  git remote set-url ado "$ADO_URL"
else
  git remote add ado "$ADO_URL"
fi

echo "remotes:"
git remote -v
echo
echo "Push to both with: ./scripts/push-mirrors.sh main"
