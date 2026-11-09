@echo off
setlocal
set "LINACCESS_MACRO_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if exist "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" set "LINACCESS_MACRO_PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
"%LINACCESS_MACRO_PS%" -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Start-Makroexport.ps1"
pause
