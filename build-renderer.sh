#!/usr/bin/env bash
# Rebuild pptx.js (the renderer the content script loads) from pptx.src.js.
# Needs `npm install` once: the source imports jszip, uuid, echarts and tslib.
set -euo pipefail
cd "$(dirname "$0")"

ESB=./node_modules/.bin/esbuild
if [ ! -x "$ESB" ]; then
  echo "no local esbuild - run 'npm install' first" >&2
  exit 1
fi

"$ESB" src/pptx.src.js --bundle --format=iife --global-name=pptxPreview \
  --target=es2020 --outfile=dist/pptx.js

# a content script cannot rely on a top level var reaching its sandbox global
printf '\nglobalThis.pptxPreview = pptxPreview;\n' >> dist/pptx.js
echo "dist/pptx.js <- src/pptx.src.js"
