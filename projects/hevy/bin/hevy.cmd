@echo off
REM Goes through the venv's python.exe, not the hevy.exe pip generates. That
REM launcher is unsigned -- it is pip's distlib stub with the entry point
REM appended as a zip, so its hash is unique to this install and can never
REM carry a signature -- and this machine enforces WDAC, which blocks it.
REM python.exe is signed by the Python Software Foundation and runs fine.
REM hevy_mcp\__main__.py calls cli.main, the same target as the `hevy` entry
REM point in pyproject.toml. -P keeps the current directory off sys.path.
"%~dp0..\.venv\Scripts\python.exe" -P -m hevy_mcp %*
exit /b %ERRORLEVEL%
