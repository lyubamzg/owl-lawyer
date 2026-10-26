#!/bin/bash
# ШАГ 2: сбор готовых документов из книги ввода
cd "$(dirname "$0")" || exit 1
BOOK="housing-input.xlsx"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  ШАГ 2: СБОР ДОКУМЕНТОВ ИЗ ТАБЛИЦЫ"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# предупреждение: таблица должна быть сохранена
if ls .~lock*"$(basename "$BOOK")"# >/dev/null 2>&1; then
  echo
  echo "!! Таблица открыта в LibreOffice. Читаю СОХРАНЁННУЮ версию."
  echo "!! Если правки не сохранены (Ctrl+S) — они не попадут в документы."
  echo
fi

python3 generate.py "$BOOK"

echo
echo "──────────── ОТЧЁТ-СВЕРКИ ────────────"
cat "validation-report.txt" 2>/dev/null
echo "───────────────────────────────────────"
echo
echo "Enter — открыть папку с документами (или окно закроется)."
read -r -p "[Enter] " x
OUT=$(ls -td output/202* 2>/dev/null | head -1)
if [ -n "$OUT" ] && [ -d "$OUT" ]; then
  xdg-open "$OUT"
fi