#!/usr/bin/env bash
# Markdown dosyalarındaki linkleri kontrol eder.
# Kullanım: proje kökünden ./scripts/check-links.sh veya npm run check-links

set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

if ! command -v npx &>/dev/null; then
  echo "Uyarı: npx bulunamadı. 'npm install' çalıştırıp tekrar deneyin."
  exit 1
fi

# Önce: npm install (markdown-link-check için)
if [ ! -d "node_modules/markdown-link-check" ]; then
  echo "Önce 'npm install' çalıştırın (markdown-link-check yüklenecek)."
  exit 1
fi

FAILED=0
while IFS= read -r -d '' f; do
  echo "Kontrol: $f"
  if ! npx markdown-link-check "$f" --config scripts/mlc-config.json; then
    FAILED=1
  fi
done < <(find . -name "*.md" -not -path "./node_modules/*" -print0)

exit $FAILED
