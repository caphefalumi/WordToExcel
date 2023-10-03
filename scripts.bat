@echo off
setlocal enabledelayedexpansion

rem Define the Python version to install
set PYTHON_VERSION=3.9.7

rem Define the list of Python libraries to install
set LIBRARIES=pandas python-docx regex tk pywin32 tkinter

rem Install Python
echo Installing Python %PYTHON_VERSION%...
choco install python --version=%PYTHON_VERSION% --yes

rem Verify Python installation
python --version

rem Install Python libraries
echo Installing Python libraries...
pip install --upgrade pip
for %%i in (%LIBRARIES%) do (
    echo Installing %%i...
    pip install %%i
)

echo Installation completed.

rem Pause to keep the command prompt open
pause
