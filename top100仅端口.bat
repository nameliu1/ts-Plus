@echo off
cd /d "%~dp0"

python 2.py
if errorlevel 1 goto end

python ppp.py

:end
pause