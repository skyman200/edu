#!/bin/bash
# Deploy index.html + the /api/ai Pages Function to Cloudflare Pages (kimcook-edu).
# wrangler reads functions/ from the CURRENT directory (not from the assets folder), so this builds a clean staging dir:
# only the public files, never README, .git, ._* files or a local .dev.vars.
set -euo pipefail
cd "$(dirname "$0")"
find functions -name '._*' -delete
node gas/build.mjs
node vibe/build.mjs
S=$(mktemp -d /tmp/kimcook-edu-deploy.XXXXXX)
trap 'rm -rf "$S"' EXIT
mkdir -p "$S/public"
cp index.html live.html check.html gas.html kit.html ta.html studio.html vibe-kit.js doc-reader.js gas-diag.js gas-recipes.js ping.js ai-class-slides.pdf ppt-kit.js office-kit.js "$S/public/"
cp -R vendor "$S/public/vendor"
find "$S" -name '._*' -delete
cp -R functions "$S/functions"
rm -f "$S/functions/_lib/"*.test.mjs
cd "$S"
npx --yes wrangler@latest pages deploy public --project-name kimcook-edu --branch main --commit-dirty=true
