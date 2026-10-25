@echo off
setlocal
cd /d "%~dp0"
if errorlevel 1 goto directory_error
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
if not exist ".venv\Scripts\python.exe" goto missing_python
echo EstimateM: http://127.0.0.1:8001
echo Keep this window open. Press Ctrl+C to stop.
".venv\Scripts\python.exe" run.py
set "ESTIMATEM_EXIT=%ERRORLEVEL%"
if "%ESTIMATEM_EXIT%"=="0" exit /b 0
echo.
echo Startup failed. Check the error above and DEPLOY_WINDOWS_OFFLINE.md.
pause
exit /b %ESTIMATEM_EXIT%

:missing_python
echo Missing .venv\Scripts\python.exe.
echo Complete the offline installation in DEPLOY_WINDOWS_OFFLINE.md first.
pause
exit /b 1

:directory_error
echo Cannot open the application directory.
pause
exit /b 1
