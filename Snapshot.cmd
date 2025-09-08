@echo off
setlocal
set "PS7=%ProgramFiles%\PowerShell\7\pwsh.exe"
set "SCRIPT=%~dp0Tools\Snapshot_rk.ps1"

if exist "%PS7%" (
  "%PS7%" -NoLogo -ExecutionPolicy Bypass -File "%SCRIPT%" %*
) else (
  powershell -NoLogo -ExecutionPolicy Bypass -File "%SCRIPT%" %*
)
