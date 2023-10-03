@echo off
setlocal enabledelayedexpansion

rem Define the Python version to install
set PYTHON_VERSION=3.9.7

rem Define the list of Python libraries to install
set LIBRARIES=pandas python-docx regex tk pywin32

rem Check if the script is in the "rewrite" state
if not defined REWRITE (
    echo Initial setup...
    
    rem Install Chocolatey
    echo Installing Chocolatey...
    @"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "[System.Net.ServicePointManager]::SecurityProtocol = 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))" && SET "PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin"

    rem Install Python
    echo Installing Python %PYTHON_VERSION%...
    choco install python --version=%PYTHON_VERSION% --yes

    rem Update the PATH to include Python and Scripts directory
    SETX PATH "%PATH%;%PROGRAMDATA%\chocolatey\lib\python\tools\;%PROGRAMDATA%\chocolatey\lib\python\tools\Scripts\"

    rem Verify Python installation
    python --version

    rem Install Python libraries
    echo Installing Python libraries...
    for %%i in (%LIBRARIES%) do (
        echo Installing %%i...
        pip install %%i
    )

    rem Set the REWRITE environment variable to mark the script as processed
    setx REWRITE 1

    rem Rewrite the script to only contain "python gui.py"
    echo @echo off > %~dp0rewrite.bat
    echo python gui.py >> %~dp0rewrite.bat
    echo Script has been rewritten. Please run the new script.
    pause
    exit
)

rem The script is in the "rewrite" state, so just run gui.py
python gui.py
