@echo off
setlocal
set "ROOT=%~dp0"
set "PS7=%ProgramFiles%\PowerShell\7\pwsh.exe"
if exist "%PS7%" (set "PS=%PS7%") else set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT%Tools\AI-Brief.ps1"