@echo off
setlocal
rem --- PowerShell 7 si dispo, sinon Windows PowerShell ---
set "ROOT=%~dp0"
set "PS7=%ProgramFiles%\PowerShell\7\pwsh.exe"
if exist "%PS7%" (set "PS=%PS7%") else set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

rem --- Lance l’instantané complet (CSV + manifest/brief/diff + docs si dispo) ---
"%PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%ROOT%Tools\snapshot_rk.ps1"
