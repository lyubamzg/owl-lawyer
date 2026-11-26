#!/bin/bash
# Сборка «СоваЮрист-пакет.zip» из содержимого репозитория.
#
# В архив попадает только то, что нужно пользователю: программа, шаблоны
# документов, база почтовых индексов, установщики и демо-пример.
# Рабочие данные юриста (книги ввода, сканы, output*) в архив не идут.
#
# Запуск: bash packaging/build-package.sh
set -eu

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$(mktemp -d)"
DEST="$WORK/СоваЮрист-пакет"
trap 'rm -rf "$WORK"' EXIT

mkdir -p "$DEST/app"

# --- верхний уровень ---
cp -f "$ROOT/README.txt" "$DEST/"
cp -f "$ROOT/install-linux.sh" "$ROOT/install-windows.bat" "$ROOT/requirements.txt" \
      "$ROOT/verification-checklist.txt" "$ROOT/windows-downloads.txt" \
      "$ROOT/windows-installation.txt" "$DEST/"
cp -r "$ROOT/icons" "$ROOT/templates" "$ROOT/incoming" "$DEST/"

# --- программа ---
find "$ROOT/app" -maxdepth 1 -type f \
     \( -name '*.py' -o -name '*.sh' -o -name '*.bat' -o -name '*.json' \) \
     -exec cp -f {} "$DEST/app/" \;
for name in input-workbook.xlsx demo-housing.xlsx \
            demo-family-questionnaire.xlsx demo-consumer-questionnaire.xlsx; do
    [ -f "$ROOT/app/$name" ] && cp -f "$ROOT/app/$name" "$DEST/app/"
done
cp -r "$ROOT/app/icons" "$ROOT/app/pretension-templates" "$DEST/app/" 2>/dev/null || true
# база почтовых индексов РФ (публичные данные ОПС «Почты России»)
mkdir -p "$DEST/app/data"
cp -f "$ROOT/app/data/"postal_*.gz "$DEST/app/data/" 2>/dev/null || true

# --- архив ---
cd "$WORK"
rm -f "$ROOT/release/СоваЮрист-пакет.zip"
zip -q -r -9 "$ROOT/release/СоваЮрист-пакет.zip" "СоваЮрист-пакет"
cp -f "$ROOT/release/СоваЮрист-пакет.zip" "$HOME/Desktop/СоваЮрист-пакет.zip"

echo "Готово:"
echo "  release/СоваЮрист-пакет.zip"
echo "  ~/Desktop/СоваЮрист-пакет.zip"
unzip -l "$ROOT/release/СоваЮрист-пакет.zip" | tail -3
