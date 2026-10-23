@echo off
setlocal
title AutoGrammar - install for Word and Excel
rem Registers AutoGrammar with Word and Excel for this Windows user (no admin rights needed).
rem Office loads every manifest listed under HKCU\...\WEF\Developer when Word/Excel start.

set "SRC=%~dp0manifest-word-excel.xml"
if not exist "%SRC%" (
  echo Could not find manifest-word-excel.xml next to this file.
  echo Please unzip AutoGrammar-Office.zip first and run this file from the office-addin folder.
  pause
  exit /b 1
)

set "DEST=%LOCALAPPDATA%\AutoGrammar"
if not exist "%DEST%" mkdir "%DEST%"
copy /Y "%SRC%" "%DEST%\manifest-word-excel.xml" >nul

set "MANIFEST=%DEST%\manifest-word-excel.xml"
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%MANIFEST%" /t REG_SZ /d "%MANIFEST%" /f >nul
if %errorlevel% neq 0 (
  echo Could not register AutoGrammar. Please send a screenshot of this window.
  pause
  exit /b 1
)

rem Remove the older shared-folder setup if a previous version of this installer added it.
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{4E1B7A2C-6F3D-4C8B-9A51-2D7E0F6B3C19}" /f >nul 2>&1

echo.
echo  AutoGrammar is installed for Word and Excel.
echo.
echo  Next:
echo   1. Close ALL Word and Excel windows, then open Word again.
echo   2. On the Home tab, at the right end, click AutoGrammar.
echo      (If Word asks whether to trust or load the add-in, click Yes / Trust.)
echo   3. In the panel: Settings ^> paste your API key ^> Save.
echo   Excel works the same way.
echo.
echo  To uninstall later, run Uninstall-Word-Excel.bat.
echo.
pause
