@echo off
REM ================================================================
REM  build_bundle.bat  —  Run this ONCE on your laptop before
REM  building the installer in Inno Setup.
REM
REM  Output: bundle\ folder with everything self-contained
REM  Then open installer.iss in Inno Setup and press F9 to compile.
REM ================================================================

setlocal
set BUNDLE=bundle
set PY_VER=3.12.10
set PY_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip
set PIP_URL=https://bootstrap.pypa.io/get-pip.py

echo.
echo [1/6] Cleaning old bundle...
if exist %BUNDLE% rmdir /s /q %BUNDLE%
mkdir %BUNDLE%

echo [2/6] Downloading Python %PY_VER% embeddable...
mkdir %BUNDLE%\python
curl -L -o %BUNDLE%\python.zip %PY_URL%
tar -xf %BUNDLE%\python.zip -C %BUNDLE%\python
del %BUNDLE%\python.zip

echo [3/6] Enabling pip in embedded Python...
powershell -Command "(Get-Content '%BUNDLE%\python\python312._pth') -replace '#import site','import site' | Set-Content '%BUNDLE%\python\python312._pth'"
curl -L -o %BUNDLE%\python\get-pip.py %PIP_URL%
%BUNDLE%\python\python.exe %BUNDLE%\python\get-pip.py --no-warn-script-location
del %BUNDLE%\python\get-pip.py

echo [4/6] Installing packages into bundle\python...
%BUNDLE%\python\python.exe -m pip install ^
    watchdog openpyxl pdf2image Pillow numpy ^
    pytesseract pywin32 ^
    --target %BUNDLE%\python\Lib\site-packages ^
    --no-warn-script-location -q
echo     Packages installed.

echo [5/6] Copying Tesseract and Poppler...
xcopy /E /I /Q "C:\Users\aayan.boradia\Downloads\Tesseract-OCR" "%BUNDLE%\Tesseract-OCR\"
mkdir %BUNDLE%\poppler
xcopy /E /I /Q "C:\Users\aayan.boradia\Downloads\poppler-26.02.0\Library\bin" "%BUNDLE%\poppler\bin\"
echo     Tesseract and Poppler copied.

echo [6/6] Copying bot source files...
for %%f in (
    app.py
    updater.py
    processor.py
    parser.py
    extractor.py
    excel_writer.py
    mailer.py
    watcher.py
    cleanup.py
) do (
    if exist %%f (
        copy /Y %%f %BUNDLE%\%%f > nul
        echo     Copied %%f
    ) else (
        echo     WARNING: %%f not found, skipping
    )
)

REM Write version file
echo 1.0.0> %BUNDLE%\version.txt

REM Write default empty config
echo {}> %BUNDLE%\config.json

echo.
echo [PyInstaller] Building ShipmentBot.exe...
py -m pip install pyinstaller -q
py -m PyInstaller ^
    --name ShipmentBot ^
    --onefile ^
    --windowed ^
    --distpath %BUNDLE% ^
    --workpath build_tmp ^
    --specpath build_tmp ^
    app.py
rmdir /s /q build_tmp 2>nul
del /q ShipmentBot.spec 2>nul

echo.
echo ================================================================
echo  Bundle ready in:  bundle\
echo  Now open installer.iss in Inno Setup and press F9 to compile.
echo  Final installer will be at:  dist\ShipmentBotSetup.exe
echo ================================================================
pause