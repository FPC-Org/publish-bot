@echo off
echo === FPC Preview — EXE Builder ===
echo.

echo Checking for PyInstaller...
python -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
    echo.
)

echo Building FPC_Preview.exe...
echo.

python -m PyInstaller preview-build\preview.spec ^
    --distpath preview-build\dist ^
    --workpath preview-build\build ^
    --noconfirm

echo.
if exist "preview-build\dist\FPC_Preview.exe" (
    echo Build complete.
    echo Output: preview-build\dist\FPC_Preview.exe
) else (
    echo Build may have failed — check output above.
)

echo.
pause
