@echo on
setlocal EnableExtensions EnableDelayedExpansion
set "SCRIPT_DIR=%~dp0"
set "ROOT=%SCRIPT_DIR%.."
pushd "%ROOT%"
set "PY=py -3.13"
%PY% -m pip install --user -U pip
%PY% -m pip install --user pyinstaller "comtypes>=1.4.8" pycaw==20240210 pywin32==311 psutil Pillow pystray
call "%SCRIPT_DIR%clean_comtypes_cache.bat"
%PY% -m PyInstaller --name "GalBGMController-CLI" --icon "assets\app.ico" --onefile --console --additional-hooks-dir "%SCRIPT_DIR%" --version-file "%SCRIPT_DIR%version_info.txt" --clean app_cli.py
echo.
echo ===== Build done =====
echo   dist\GalBGMController-CLI.exe
pause
popd
endlocal
