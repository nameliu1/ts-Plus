@echo off
cd /d "%~dp0"
echo Starting the process...

:: 删除本地目录下的port.txt
if exist port.txt del /F /Q port.txt
if exist url.txt del /F /Q url.txt

:: Step 1: Synchronously run the Python script 2.py
python 2.py
if errorlevel 1 goto end

python ppp.py
if errorlevel 1 goto end

if exist res.json del /F /Q res.json
if exist res_processed.txt del /F /Q res_processed.txt
if exist res_processed.xlsx del /F /Q res_processed.xlsx

python 1.py
if errorlevel 1 goto end

:end
pause