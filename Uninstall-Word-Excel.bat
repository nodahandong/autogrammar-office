@echo off
title AutoGrammar - uninstall from Word and Excel
set "MANIFEST=%LOCALAPPDATA%\AutoGrammar\manifest-word-excel.xml"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%MANIFEST%" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{4E1B7A2C-6F3D-4C8B-9A51-2D7E0F6B3C19}" /f >nul 2>&1
echo AutoGrammar was removed from Word and Excel. Restart Word and Excel.
pause
