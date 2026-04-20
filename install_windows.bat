@echo off
REM Install script for TTTurnier KO Reorder on Windows
REM This script sets up a Python virtual environment and installs dependencies

echo Setting up Python virtual environment...
python -m venv venv

echo Activating virtual environment...
call venv\Scripts\activate

echo Installing dependencies from requirements.txt...
pip install -r requirements.txt

echo Installation complete!
echo To run the script:
echo   call venv\Scripts\activate
echo   python TTTurnier_KO_Reorder.py [-v] [-n] database.mdb
