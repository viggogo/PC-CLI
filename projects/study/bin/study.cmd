@echo off
REM PowerShell resolves .ps1 from PATH ahead of PATHEXT, so a sibling study.ps1
REM would shadow this shim and get blocked by the machine's execution policy.
REM This file therefore lives alone in bin\ -- bin\ is what goes on PATH, not the
REM project folder, which also holds install.ps1 and test.ps1 (generic names that
REM would otherwise become global commands in every terminal).
REM -ExecutionPolicy Bypass applies to THIS invocation only, nothing system-wide.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\study.ps1" %*
exit /b %ERRORLEVEL%
