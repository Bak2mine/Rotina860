@echo off
REM Rebuild Rotina860.exe with cleanup

echo Cleaning old build artifacts...
if exist dist rmdir /s /q dist
if exist build rmdir /s /q build

echo Building new executable...
python build_exe.py

echo.
echo Done! New Rotina860.exe is ready in the dist folder.
pause
