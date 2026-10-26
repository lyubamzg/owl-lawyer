@echo off
chcp 65001 >nul
setlocal
set PYTHONUTF8=1
set "SRC=%~dp0"
set "DEST=%USERPROFILE%\Desktop\OwlLawyer"
set "POD=%USERPROFILE%\Desktop\filing"

echo ════════════════════════════════════════════════
echo   УСТАНОВКА «СОВА-ЮРИСТ» (Windows)
echo ════════════════════════════════════════════════
echo Папка программы: %DEST%
echo.
echo ── Проверка 1/5: Python ────────────────────────
where python >nul 2>&1
if errorlevel 1 (
  echo   НЕТ Python. Скачайте с https://www.python.org/downloads/
  echo   В самом первом окне установщика поставьте ГАЛОЧКУ
  echo   "Add python.exe to PATH" и вторую "tcl/tk and IDLE"!
  goto :advice
)
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
  echo   Python есть, но БЕЗ tkinter (окно приложения).
  echo   Переустановите Python с галочкой "tcl/tk and IDLE".
  goto :advice
)
echo   ОК.

echo ── Проверка 2/5: библиотеки openpyxl и python-docx ──
python -m pip install --upgrade pip >nul 2>&1
python -m pip install openpyxl python-docx
if errorlevel 1 goto :advice
echo   ОК.

echo ── Проверка 3/5: Tesseract OCR (русский) ──────
where tesseract >nul 2>&1
if errorlevel 1 (
  echo   НЕТ Tesseract (распознавание сканов Ф-9).
  echo   Установите tesseract-ocr-w64 по ссылке из windows-downloads.txt.
  echo   ВАЖНО: при установке отметьте языковой пакет RUSSIAN!
  goto :advice
)
echo   ОК.

echo ── Проверка 4/5: Poppler (pdftoppm) ────────────
where pdftoppm >nul 2>&1
if errorlevel 1 (
  echo   НЕТ Poppler - он режет PDF на страницы для распознавания.
  echo   Скачайте poppler-windows (см. windows-downloads.txt), распакуйте
  echo   в C:\poppler и добавьте папку C:\poppler\...\bin в PATH -
  echo   в windows-installation.txt это расписано по кликам.
  goto :advice
)
echo   ОК.

echo ── Проверка 5/5: LibreOffice ───────────────────
set SOFF=
if exist "C:\Program Files\LibreOffice\program\soffice.exe" set "SOFF=C:\Program Files\LibreOffice\program\soffice.exe"
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" set "SOFF=C:\Program Files (x86)\LibreOffice\program\soffice.exe"
where soffice >nul 2>&1
if errorlevel 1 if not defined SOFF (
  echo   НЕТ LibreOffice (просмотр и печать таблиц и документов).
  echo   Установите с https://ru.libreoffice.org/ - файл .msi, "далее-далее".
  goto :advice
)
echo   ОК.

echo ── Копирование файлов ──────────────────────────
xcopy /e /i /y "%SRC%app\*" "%DEST%\" >nul
xcopy /e /i /y "%SRC%templates\*"  "%POD%\"   >nul
mkdir "%POD%\incoming" 2>nul
echo   Программа:    %DEST%
echo   Шаблоны .doc: %POD%

echo ── Ярлыки на рабочем столе ─────────────────────
powershell -NoProfile -Command "$ws=(New-Object -ComObject WScript.Shell); $py=((Get-Command pythonw -ErrorAction SilentlyContinue).Source); if(-not $py){$py=((Get-Command python).Source)}; $s=$ws.CreateShortcut($env:USERPROFILE+'\Desktop\OwlLawyer.lnk'); $s.TargetPath=$py; $s.Arguments='%DEST%\desktop_app.py'; $s.WorkingDirectory='%DEST%'; $s.IconLocation='%DEST%\icons\owl-lawyer.png'; $s.Save(); foreach($p in @(@('1-import-data','step1-import.bat'),@('2-generate-documents','step2-generate.bat'))){ $s=$ws.CreateShortcut($env:USERPROFILE+'\Desktop\'+$p[0]+'.lnk'); $s.TargetPath='%ComSpec%'; $s.Arguments='/c \"'+$p[1]+'\"'; $s.WorkingDirectory='%DEST%'; $s.Save() }" >nul 2>&1
echo   Готово: 3 ярлыка (если пусто - создайте вручную, как в инструкции).

echo ── Проверка: демо-прогон ───────────────────────
pushd "%DEST%"
python generate.py "demo-housing.xlsx"
if errorlevel 1 ( echo !!! демо-прогон упал - см. windows-installation.txt & popd & goto :advice )
popd
echo   ОК: демо-документы в %DEST%\output (учебные, можно удалить).

echo.
echo ════════════════════════════════════════════════
echo   УСТАНОВКА ЗАВЕРШЕНА
echo ════════════════════════════════════════════════
start "" pythonw "%DEST%\desktop_app.py"
echo Открылось окно «Сова-юрист». Дальше - verification-checklist.txt
pause
exit /b 0

:advice
echo.
echo ────────────────────────────────────────────────
echo Установите недостающее по файлу windows-downloads.txt
echo (или по шагам из windows-installation.txt), затем
echo запустите install-windows.bat ещё раз.
echo ────────────────────────────────────────────────
pause
exit /b 1