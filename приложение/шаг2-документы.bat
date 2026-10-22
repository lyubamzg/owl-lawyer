@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
echo ================================================
echo   ШАГ 2: СБОР ДОКУМЕНТОВ ИЗ ТАБЛИЦЫ
echo ================================================
echo.
echo ВАЖНО: таблица должна быть СОХРАНЕНА (Ctrl+S) в LibreOffice!
set /p GO=Нажмите Enter для запуска (или Ctrl+C для отмены)...
python generate.py "Книга-ввода (портянка).xlsx"
if errorlevel 1 (
  python generate.py "Книга-ввода.xlsx"
  if errorlevel 1 (
    echo.
    echo !!! Сбор завершился с ошибкой - прочитайте отчёт выше.
    pause
    exit /b 1
  )
)
echo.
echo ────────── ОТЧЁТ-СВЕРКИ ──────────
type "Отчёт-сверки.txt"
echo ──────────────────────────────────
echo.
set OUT=
for /f delims= %%i in ('dir /b /ad /on "Выход\202*"') do set "OUT=%%i"
if defined OUT start "" explorer "Выход\%OUT%"
echo Папка с документами открыта.
pause