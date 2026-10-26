@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"
echo ================================================
echo   ШАГ 1: ЗАГРУЗКА ДАННЫХ (Портянка + Ф-9)
echo ================================================
echo.
echo Ожидаю файлы в папке ..\filing\incoming\ :
echo   - Портянка*.xlsx  (расчёт от бухгалтерии)
echo   - Ф-9*.pdf        (сканы справок о регистрации)
echo.
echo Первый прогон: создаёт таблицу и (если нужно) f9-index.xlsx
python import_housing.py
if errorlevel 1 (
  echo.
  echo !!! Импорт завершился с ошибкой - прочитайте сообщение выше.
  pause
  exit /b 1
)
python -c "import os;print(' индекс создан' if os.path.exists('f9-index.xlsx') else ' индекс не нужен - пропустите дальше')"
echo.
echo Если выше сказано «индекс создан»: откройте «f9-index.xlsx»,
echo впишите ЛС напротив номеров страниц, СОХРАНИТЕ (Ctrl+S),
echo и только после этого нажимайте Enter - будет второй прогон.
set /p GO=Нажмите Enter, чтобы сделать второй прогон (или Ctrl+C для отмены)...
python import_housing.py
if errorlevel 1 (
  echo !!! Второй прогон завершился с ошибкой.
  pause
  exit /b 1
)
echo.
echo Готово: создана «housing-input.xlsx».
set SOFF=C:\Program Files\LibreOffice\program\soffice.exe
if exist "%SOFF%" (start "" "%SOFF%" "%CD%\housing-input.xlsx") else (start "" soffice "%CD%\housing-input.xlsx")
echo Таблица открыта в редакторе. Дочистите лист «Сверка импорта»,
echo СОХРАНИТЕ и нажмите «2 Собрать документы».
pause