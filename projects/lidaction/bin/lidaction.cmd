@echo off
REM PowerShell resolves .ps1 from PATH ahead of PATHEXT, so a sibling lidaction.ps1
REM would shadow this shim and get blocked by the machine's execution policy.
REM This file therefore lives alone in bin\ -- bin\ is what goes on PATH, not the
REM project folder. Keeping the .ps1 out of PATH is what makes the command work.
REM -ExecutionPolicy Bypass applies to THIS invocation only, nothing system-wide.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\lidaction.ps1" %*
exit /b %ERRORLEVEL%
