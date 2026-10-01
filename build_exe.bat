@echo off
setlocal
cd /d "%~dp0"
echo [1/3] Installing build packages...
py -m pip install --upgrade pip
if errorlevel 1 goto FAIL
py -m pip install requests beautifulsoup4 openpyxl pyinstaller
if errorlevel 1 goto FAIL
echo [2/3] Building Windows EXE...
py -m PyInstaller --noconfirm --clean --onefile --windowed --name AsanLibraryFinder AsanLibraryFinder.py
if errorlevel 1 goto FAIL
echo [3/3] Done.
echo EXE: %cd%\dist\AsanLibraryFinder.exe
pause
exit /b 0
:FAIL
echo.
echo BUILD FAILED. Check the error above.
pause
exit /b 1
