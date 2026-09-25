@echo off
setlocal
pushd "%~dp0"
set "PY=py -3.13"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
%PY% app_cli.py %*
set "RESULT=%ERRORLEVEL%"
popd
exit /b %RESULT%
