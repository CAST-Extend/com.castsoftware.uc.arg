@echo off
setlocal EnableDelayedExpansion

:: Initialize the default PIP install version
set PIP_INSTALL_VERSION=1.6.9

:: Set Python version
set "python_version=3.10.6"

:: Check script name
if "%~0"=="install.bat" (

    :: Check for command-line arguments
    if not "%1"=="" (
        if "%1"=="upgrade" (
            echo Detected upgrade command...

            if not "%2"=="" (
                set "PIP_INSTALL_VERSION=%2"
             
            ) else (
                echo ...
            )
        )
    ) else (
        echo ...
    )
)
echo .
echo .
echo .
echo .
echo ======================================================================
echo Installation version: com.castsoftware.uc.arg == %PIP_INSTALL_VERSION% 
echo ======================================================================
echo .
echo .
echo .
echo .


:: Check for administrative permissions
echo Administrative permissions required. Detecting permissions...
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Success: Administrative permissions confirmed.
    goto PythonInstallCheck
) else (
    echo Failure: Current permissions inadequate.
    echo Please run the script as an Administrator.
    pause
    exit
)

:PythonInstallCheck
python --version >nul 2>&1
if %errorLevel% == 0 (
    echo Python already installed.
    goto PipCheck
) else (
    goto PythonInstall
)

:PythonInstall
:: Define the URL and file name of the Python installer
set "url=https://www.python.org/ftp/python/%python_version%/python-%python_version%-amd64.exe"
set "installer=python-%python_version%-amd64.exe"

:: Define the installation directory
set "targetdir1=C:\Python%python_version%"

:: Download the Python installer
powershell -Command "(New-Object Net.WebClient).DownloadFile('%url%', '%~dp0%installer%')"

:: Install Python
echo Installing Python Version %python_version%...
start /wait "" "%~dp0%installer%" /quiet /passive TargetDir="%targetdir1%" Include_test=0 PrependPath=1 ^
&& (echo Done.) || (echo Failed!)
echo.

:: Cleanup
del "%~dp0%installer%"

goto PipCheck

:PipCheck

python -m ensurepip --default-pip
call pip --version 2>nul
if errorlevel 1 (
  echo PIP not found, please ensure it's installed.
  pause
  exit
) else (
  goto GetDriveLetter
)



:GetDriveLetter
set /p drive=Please enter the path where you want to install ARG: 
set installPath=%drive%\ARG

if not exist "%installPath%" (
    mkdir "%installPath%"
    cd "%installPath%"
    echo ARG folder created at %installPath%
) else (
    cd "%installPath%"
    echo ARG folder already exists at %installPath%. Proceeding with installation.
)

set CODE_FOLDER=%installPath%
goto VenvSetup

:VenvSetup
echo creating virtual environment 
python -m venv "%CODE_FOLDER%\.venv"

:: Rename .venv if it exists
if exist "%CODE_FOLDER%\.venv" (
    for /f "delims=" %%a in ('wmic os get LocalDateTime ^| find "."') do (
            set datetime=%%a
        )
        set datetime=!datetime:~0,14!
        set datetime=!datetime:~0,4!!datetime:~4,2!!datetime:~6,2!_!datetime:~8,2!!datetime:~10,2!!datetime:~12,2!
        rename "%CODE_FOLDER%\.venv" "venv_!datetime!"
    
)

:: Create a new .venv folder
if not exist "%CODE_FOLDER%\.venv" (
    python -m venv "%CODE_FOLDER%\.venv"
    if errorlevel 1 goto VenvFail
)

:: Copying essential files to the destination folder
copy "%~dp0arg.bat" "%installPath%"
copy "%~dp0cause.json" "%installPath%"
copy "%~dp0config.json" "%installPath%"
copy "%~dp0README.md" "%installPath%"
copy "%~dp0*.pptx" "%installPath%"

:: Activate the virtual environment
call "%CODE_FOLDER%\.venv\Scripts\activate"

pip install com.castsoftware.uc.arg==%PIP_INSTALL_VERSION%

goto End

:VenvFail
echo Unable to install virtual environment
goto Usage

:Usage
echo Usage: install.bat
pause
exit /b

:End
echo Script completed!
pause
exit
