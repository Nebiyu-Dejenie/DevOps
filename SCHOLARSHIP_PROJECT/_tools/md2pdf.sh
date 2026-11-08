#!/bin/bash
# usage: md2pdf.sh file.md [landscape]  -> file.pdf next to the md (HTML via pandoc, printed by Chromium/Playwright)
set -e
SPD="$(cd "$(dirname "$(realpath "$0")")" && pwd)"
md="$1"; base="${md%.md}"; tmp=$(mktemp -d)
pandoc "$md" -s --metadata lang=en --css "$SPD/doc.css" --embed-resources -o "$tmp/doc.html" 2>/dev/null || pandoc "$md" -s -c "$SPD/doc.css" --self-contained -o "$tmp/doc.html"
node "$SPD/md2pdf.js" "$tmp/doc.html" "$(realpath "$base").pdf" "${2:-portrait}"
rm -rf "$tmp"; echo "$(stat -c %s "$base.pdf") $base.pdf $(pdfinfo "$base.pdf" | grep Pages)"
