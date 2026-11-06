#!/bin/bash
# Deploy index.html + the /api/ai Pages Function to Cloudflare Pages (kimcook-edu).
# wrangler reads functions/ from the CURRENT directory (not from the assets folder), so this builds a clean staging dir:
# only the public files, never README, .git, ._* files or a local .dev.vars.
set -euo pipefail
cd "$(dirname "$0")"
find functions -name '._*' -delete
S=$(mktemp -d /tmp/kimcook-edu-deploy.XXXXXX)
trap 'rm -rf "$S"' EXIT
mkdir -p "$S/public"
cp index.html check.html ping.js ai-class-slides.pdf "$S/public/"
cp -R functions "$S/functions"
cd "$S"
npx --yes wrangler@latest pages deploy public --project-name kimcook-edu --branch main --commit-dirty=true
