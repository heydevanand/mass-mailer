@echo off

REM Install Python using Winget
echo Installing Python...
winget install python

REM Check if Python is installed successfully
echo Checking Python version...
python --version

REM Install pip
echo Installing pip...
python -m pip install --upgrade pip

REM Install dependencies
echo Installing dependencies...
pip install pandas openpyxl

REM Run mail.py
echo Running mail.py...
python mail.py
REM Press any key to exit
echo Press any key to exit...
pause >nul