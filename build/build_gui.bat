@echo on
setlocal EnableExtensions EnableDelayedExpansion
set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "ROOT=%%~fI"
pushd "%ROOT%"
set "PY=py -3.13"
set "SCRIPT=%ROOT%\app_gui.py"
set "ICON=%ROOT%\assets\app.ico"
set "HOOKS=%ROOT%\build"
set "VER=%ROOT%\build\version_info.txt"
if not exist "%SCRIPT%" echo ERROR: app_gui.py not found at "%SCRIPT%" & goto end
if not exist "%ICON%" echo ERROR: app icon not found at "%ICON%" & goto end
%PY% -m pip install --user -U pip
%PY% -m pip install --user pyinstaller "comtypes>=1.4.8" pycaw==20240210 pywin32==311 psutil Pillow pystray
%PY% -m PyInstaller --name "GalBGMController" --icon "%ICON%" --noconsole --add-data "%ROOT%\assets;assets" --additional-hooks-dir "%HOOKS%" --version-file "%VER%" --clean "%SCRIPT%"
%PY% -m PyInstaller --name "GalBGMController-single" --icon "%ICON%" --noconsole --onefile --add-data "%ROOT%\assets;assets" --additional-hooks-dir "%HOOKS%" --version-file "%VER%" --clean "%SCRIPT%"
echo.
echo ===== Build done =====
echo   dist\GalBGMController\GalBGMController.exe
echo   dist\GalBGMController-single.exe
:end
pause
popd
endlocal
