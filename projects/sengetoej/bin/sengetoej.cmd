@echo off
REM Goes through the venv's python.exe, not the sengetoej.exe pip generates.
REM That launcher is unsigned -- it is pip's distlib stub with the entry point
REM appended as a zip, so its hash is unique to this install and can never
REM carry a signature -- and this machine enforces WDAC, which blocks it.
REM python.exe is signed by the Python Software Foundation and runs fine.
REM -P keeps the current directory off sys.path, so a stray sengetoej.py in
REM whatever folder you are standing in cannot shadow the real package.
"%~dp0..\.venv\Scripts\python.exe" -P -m sengetoej %*
exit /b %ERRORLEVEL%
