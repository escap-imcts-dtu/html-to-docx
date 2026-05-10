#!/usr/bin/env bash
# Push a branch to BOTH GitHub and ADO mirrors.
#
# ADO requires auth. We acquire a short-lived bearer token via the Azure
# CLI (you must `az login` first) — no PAT to manage and no token to leak.
# GitHub uses whatever credential helper you already have set up (gh, ssh, …).
#
# Usage: ./scripts/push-mirrors.sh [branch]   # default: current branch

set -euo pipefail

BRANCH="${1:-$(git rev-parse --abbrev-ref HEAD)}"
ADO_RESOURCE="499b84ac-1321-427f-aa17-267ca6975798"  # Azure DevOps app id

echo "→ pushing $BRANCH to GitHub (origin)…"
git push origin "$BRANCH"

echo "→ pushing $BRANCH to ADO (ado)…"
TOKEN=$(az account get-access-token --resource "$ADO_RESOURCE" --query accessToken -o tsv 2>/dev/null \
  || { echo "az CLI not logged in. Run: az login"; exit 1; })
git -c http.extraheader="Authorization: Bearer $TOKEN" push ado "$BRANCH"

echo "✓ both mirrors updated to $(git rev-parse --short "$BRANCH")"
