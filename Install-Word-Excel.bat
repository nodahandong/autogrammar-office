@echo off
setlocal
title AutoGrammar - install for Word and Excel
rem Sets up AutoGrammar for Word and Excel on Windows:
rem copies the manifest to C:\AutoGrammar, shares that folder, and marks it as a trusted add-in catalog.

net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Asking for permission to set up the AutoGrammar folder...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)

set "SRC=%~dp0manifest-word-excel.xml"
if not exist "%SRC%" (
  echo Could not find manifest-word-excel.xml next to this file.
  echo Please unzip AutoGrammar-Office.zip first and run this file from the office-addin folder.
  pause
  exit /b 1
)

if not exist "C:\AutoGrammar" mkdir "C:\AutoGrammar"
copy /Y "%SRC%" "C:\AutoGrammar\manifest-word-excel.xml" >nul

net share AutoGrammar /delete >nul 2>&1
net share AutoGrammar="C:\AutoGrammar" /GRANT:Everyone,READ >nul
if %errorlevel% neq 0 (
  echo Could not share the folder. Your company may block this; ask IT or use Word/Excel on office.com.
  pause
  exit /b 1
)

set "KEY=HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{4E1B7A2C-6F3D-4C8B-9A51-2D7E0F6B3C19}"
reg add "%KEY%" /v Id /t REG_SZ /d "{4E1B7A2C-6F3D-4C8B-9A51-2D7E0F6B3C19}" /f >nul
reg add "%KEY%" /v Url /t REG_SZ /d "\\%COMPUTERNAME%\AutoGrammar" /f >nul
reg add "%KEY%" /v Flags /t REG_DWORD /d 1 /f >nul

echo.
echo  AutoGrammar is set up for Word and Excel.
echo.
echo  Next:
echo   1. Close Word and Excel completely, then open Word.
echo   2. Home ^> Add-ins ^> More Add-ins  (or Insert ^> My Add-ins)
echo   3. Open the SHARED FOLDER tab, choose AutoGrammar, click Add.
echo   4. Click AutoGrammar on the Home tab ^> Settings ^> paste your API key ^> Save.
echo   Do steps 2-4 in Excel too.
echo.
pause
