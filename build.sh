#!/usr/bin/env bash
# Store submission builds:
#   hover-file-preview-firefox.zip  -> addons.mozilla.org
#   hover-file-preview-chrome.zip   -> Chrome Web Store
# Both are plain zips with manifest.json at the root; the unpacked trees under
# build/ are for local testing only.
set -euo pipefail
cd "$(dirname "$0")"

VERSION=$(python3 -c "import json;print(json.load(open('package.json'))['version'])")
python3 - "$VERSION" <<'PYV'
import json, sys
m = json.load(open("manifest.json"))
if m.get("version") != sys.argv[1]:
    m["version"] = sys.argv[1]
    json.dump(m, open("manifest.json", "w"), indent=2, ensure_ascii=False)
    open("manifest.json", "a").write("\n")
    print("manifest version ->", sys.argv[1])
PYV

STAMP=$(date -u +%Y%m%d-%H%M)

rm -rf build
for t in firefox chrome; do mkdir -p "build/$t/src" "build/$t/dist" "build/$t/icons"; done

# pptx.js is the readable bundle; minify it for the package. Falls back to the
# unminified file when esbuild is not available.
if timeout 300 npx --yes esbuild@0.24.0 dist/pptx.js --minify --legal-comments=none \
     --outfile=build/firefox/dist/pptx.js >/dev/null 2>&1; then
  echo "renderer minified"
else
  cp dist/pptx.js build/firefox/dist/pptx.js
  echo "renderer copied (esbuild unavailable)"
fi
cp build/firefox/dist/pptx.js build/chrome/dist/pptx.js

# the build id goes into the packaged content.js so the card can show it
python3 - "$STAMP" <<'PY'
import io, sys
stamp = sys.argv[1]
src = io.open("src/content.js", encoding="utf8").read()
src = src.replace('const BUILD = "";', 'const BUILD = "%s";' % stamp)
for target in ("build/firefox/src/content.js", "build/chrome/src/content.js"):
    io.open(target, "w", encoding="utf8").write(src)
PY

for t in firefox chrome; do
  cp manifest.json "build/$t/"
  cp icons/icon-16.png icons/icon-32.png icons/icon-48.png icons/icon-128.png "build/$t/icons/"
  cp src/zip.js src/lib.js src/xlsx.js src/docx.js "build/$t/src/"
done

# Chrome rejects the browser_specific_settings block; Firefox needs the gecko id
python3 - <<'PY'
import json
m = json.load(open("manifest.json"))
m.pop("browser_specific_settings", None)
m["minimum_chrome_version"] = "120"
json.dump(m, open("build/chrome/manifest.json", "w"), indent=2, ensure_ascii=False)
open("build/chrome/manifest.json", "a").write("\n")
PY

rm -f hover-file-preview-firefox.zip hover-file-preview-chrome.zip
python3 - <<'PY'
import os, zipfile

FILES = [
    "manifest.json",
    "icons/icon-16.png", "icons/icon-32.png", "icons/icon-48.png", "icons/icon-128.png",
    "src/zip.js", "src/lib.js", "src/xlsx.js", "src/docx.js", "src/content.js",
    "dist/pptx.js",
]

for target, out in (
    ("firefox", "hover-file-preview-firefox.zip"),
    ("chrome", "hover-file-preview-chrome.zip"),
):
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for f in FILES:
            path = os.path.join("build", target, f)
            assert os.path.isfile(path), "missing " + path
            z.write(path, f)
    print(out, os.path.getsize(out), "bytes")
PY

# what a store reviewer would check first
python3 - <<'PY'
import json, sys, zipfile

EXPECTED = {
    "manifest.json",
    "icons/icon-16.png", "icons/icon-32.png", "icons/icon-48.png", "icons/icon-128.png",
    "src/zip.js", "src/lib.js", "src/xlsx.js", "src/docx.js", "src/content.js",
    "dist/pptx.js",
}
ok = True
for name in ("hover-file-preview-firefox.zip", "hover-file-preview-chrome.zip"):
    z = zipfile.ZipFile(name)
    names = set(z.namelist())
    m = json.loads(z.read("manifest.json"))
    problems = []
    if names != EXPECTED:
        problems.append("unexpected file set: " + str(sorted(names ^ EXPECTED)))
    if m.get("manifest_version") != 3:
        problems.append("not MV3")
    for size, path in m.get("icons", {}).items():
        if path not in names:
            problems.append("icon %s missing" % path)
    if name.endswith("chrome.zip"):
        if "browser_specific_settings" in m:
            problems.append("chrome manifest carries gecko settings")
    else:
        gecko = m.get("browser_specific_settings", {}).get("gecko", {})
        if not gecko.get("id"):
            problems.append("firefox manifest has no gecko id")
    print(("OK   " if not problems else "FAIL ") + name, "v" + m["version"], "|", len(names), "files",
          "|", round(sum(z.getinfo(n).file_size for n in names) / 1024), "KB unpacked")
    for p in problems:
        ok = False
        print("   -", p)
sys.exit(0 if ok else 1)
PY
echo "build $STAMP ready (submission zips at the repo root)"
