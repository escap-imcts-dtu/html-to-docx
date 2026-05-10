#!/usr/bin/env bash
# Configure `origin` to push to BOTH the GitHub mirror and the ADO mirror.
# Run this once after cloning so `git push origin main` writes to both.
#
# We mirror to two remotes for resilience: if one is accidentally deleted
# or becomes inaccessible, the other holds the full history.
#
# - GitHub: https://github.com/escap-imcts-dtu/html-to-docx (consumed by ESCAP)
# - ADO:    https://unescap.visualstudio.com/ESCAP-Document-Center/_git/html-to-docx (backup)

set -euo pipefail

GH_URL="https://github.com/escap-imcts-dtu/html-to-docx.git"
ADO_URL="https://unescap.visualstudio.com/ESCAP-Document-Center/_git/html-to-docx"

git remote set-url --push origin "$GH_URL"
git remote set-url --add --push origin "$ADO_URL"

# Add a separate `ado` remote too, so it's possible to push/pull ADO directly.
if ! git remote get-url ado > /dev/null 2>&1; then
  git remote add ado "$ADO_URL"
fi

echo "origin push URLs:"
git remote -v | grep '(push)'
echo
echo "Push to both with: ./scripts/push-mirrors.sh main"
