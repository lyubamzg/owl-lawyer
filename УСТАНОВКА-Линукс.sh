#!/bin/bash
# Установщик «Сова-юрист» для Linux (apt-дистрибутивы: Debian, Ubuntu, Mint...).
# Запуск: bash УСТАНОВКА-Линукс.sh
set -u

SRC="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/Desktop/Генератор"
POD="$HOME/Desktop/подача"

echo "════════════════════════════════════════════════"
echo "  УСТАНОВКА «СОВА-ЮРИСТ»  (Linux)"
echo "════════════════════════════════════════════════"
echo "Папка программы:  $DEST"
echo "Папка подачи:     $POD"
echo

# 1) Дистрибутив
if ! command -v apt-get >/dev/null 2>&1; then
  echo "ЭТО НЕ apt-дистрибутив. Установите вручную и запустите ещё раз:"
  echo "  python3 python3-tk python3-pip python3-openpyxl python3-docx"
  echo "  libreoffice tesseract-ocr tesseract-ocr-rus poppler-utils"
  exit 1
fi
SUDO=""
[ "$(id -u)" -ne 0 ] && SUDO="sudo"

echo "── 1/5  Установка программ (попросит пароль sudo) ──"
if [ "${SOVA_SKIP_APT:-0}" = "1" ]; then
  echo "   пропущено (SOVA_SKIP_APT=1 — пакеты уже установлены)"
else
$SUDO apt-get update -qq
$SUDO apt-get install -y \
  python3 python3-tk python3-pip \
  python3-openpyxl python3-docx \
  libreoffice tesseract-ocr tesseract-ocr-rus poppler-utils \
  || { echo "!!! apt не смог. Ставьте пакеты из сообщения выше."; exit 1; }
fi

# Если python-библиотеки не поставились из репозитория — pip:
python3 -c "import tkinter" 2>/dev/null || { echo "Нет tkinter — нужен пакет python3-tk"; exit 1; }
python3 -c "import openpyxl, docx" 2>/dev/null || {
  echo "Библиотеки openpyxl/python-docx не из apt — ставлю через pip..."
  python3 -m pip install --user openpyxl python-docx 2>/dev/null \
    || $SUDO apt-get install -y python3-openpyxl python3-docx \
    || { echo "!!! библиотеки не встали — см. README.txt"; exit 1; }
}

echo "── 2/5  Копирование файлов ──"
mkdir -p "$HOME/.local/share/icons"
cp -n "$SRC/иконки/"*.png "$HOME/.local/share/icons/" 2>/dev/null
mkdir -p "$DEST"
# Скрипты обновляем всегда (merge-режим), книги пользователя НЕ трогаем:
cp -f "$SRC/приложение/"*.py "$SRC/приложение/"*.sh "$DEST/"
chmod +x "$DEST/"*.sh "$DEST/рабочее_место.py"
cp -n "$SRC/приложение/Демо (портянка заполненная).xlsx" "$DEST/"
mkdir -p "$POD/входящие"
cp -n "$SRC/шаблоны/"*.doc "$POD/"
cp -f "$SRC/подача-входящие/README.txt" "$POD/входящие/" 2>/dev/null

echo "── 3/5  Ярлыки на рабочем столе ──"
TERM_BIN=""
for t in x-terminal-emulator gnome-terminal konsole xfce4-terminal ptyxis xterm kitty alacritty; do
  command -v "$t" >/dev/null 2>&1 && { TERM_BIN="$t"; break; }
done
case "$TERM_BIN" in
  gnome-terminal)  TERM_EXEC="$TERM_BIN -- /bin/bash" ;;
  ptyxis)          TERM_EXEC="ptyxis -- /bin/bash" ;;
  "")              TERM_EXEC="" ;;
  *)               TERM_EXEC="$TERM_BIN -e /bin/bash" ;;
esac

make_desktop() {  # $1=файл $2= название $3=Exec $4=Icon
  cat > "$HOME/Desktop/$1" <<EOF
[Desktop Entry]
Version=1.0
Type=Application
Name=$2
Comment=Сова-юрист: приказы и иски по ЖКУ
Exec=$3
Icon=$4
Terminal=false
Path=$DEST
Categories=Office;
EOF
  chmod +x "$HOME/Desktop/$1"
  cp -f "$HOME/Desktop/$1" "$HOME/.local/share/applications/" 2>/dev/null
}

ICON256="$HOME/.local/share/icons/сова-юрист-256.png"
make_desktop "Сова-юрист.desktop" "Сова-юрист" "python3 $DEST/рабочее_место.py" "$ICON256"
if [ -n "$TERM_EXEC" ]; then
  make_desktop "1 Загрузить данные.desktop"  "1 Загрузить данные"  "$TERM_EXEC \"$DEST/шаг1-загрузить.sh\""  "folder-open"
  make_desktop "2 Собрать документы.desktop" "2 Собрать документы" "$TERM_EXEC \"$DEST/шаг2-документы.sh\"" "document-print"
else
  # Терминала в системе нет — шаги запускаются напрямую, журнал в файл
  make_desktop "1 Загрузить данные.desktop"  "1 Загрузить данные"  \
    "python3 $DEST/импорт_портянки.py > $DEST/журнал-шаг1.txt 2>&1" "folder-open"
  make_desktop "2 Собрать документы.desktop" "2 Собрать документы" \
    "python3 $DEST/generate.py Книга-ввода\ \(портянка\).xlsx > $DEST/журнал-шаг2.txt 2>&1" "document-print"
  echo "   (терминал в системе не найден — шаги пишут журнал в файл)"
fi
echo "   найденный эмулятор терминала: ${TERM_BIN:-нет — запуск напрямую}"

echo "── 4/5  Проверка: сбор документов из демо-книги ──"
( cd "$DEST" && python3 generate.py "Демо (портянка заполненная).xlsx" ) \
  || { echo "!!! демо-прогон упал — см. README.txt, раздел «Если что-то не так»"; exit 1; }
echo "   ОК: демо-документы в $DEST/Выход (это учебные, можно удалить)."

echo "── 5/5  Запуск приложения ──"
nohup python3 "$DEST/рабочее_место.py" >/dev/null 2>&1 &
sleep 1
echo
echo "════ УСТАНОВКА ЗАВЕРШЕНА ════"
echo "Открылось окно «Сова-юрист» с кнопками. Если не открылось —"
echo "двойной клик по ярлыку «Сова-юрист» на рабочем столе."
echo "Подробности работы: $SRC/README.txt"