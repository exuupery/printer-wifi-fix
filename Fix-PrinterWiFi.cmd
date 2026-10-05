@echo off
rem Launcher: double-click to fix Wi-Fi printing on this PC (asks for admin via UAC).
rem First run without printer.psd1 starts the setup wizard.
rem To preview without changes run:  Fix-PrinterWiFi.cmd -WhatIf
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Fix-PrinterWiFi.ps1" %*
rem Exit code 99: the script continued in a separate administrator window.
if not %errorlevel%==99 pause
