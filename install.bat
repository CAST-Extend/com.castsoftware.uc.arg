@echo off
setlocal EnableDelayedExpansion

:: Initialize the default PIP install version
set PIP_INSTALL_VERSION=1.7.7

:: Set Python version
set "python_version=3.10.10

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
    exit /b 1
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
start /wait "" "%~dp0%installer%" /quiet /passive TargetDir="%targetdir1%" InstallAllUsers=1 Include_test=0 PrependPath=1 Include_pip=1  
if "%ERRORLEVEL%" == "0" goto Cleanup

echo Error installing Python!
pause
exit /b 1

:Cleanup
echo. 
echo Python installation completed successfully!
echo. 
path=%targetdir1%;%targetdir1%\scripts;%path%

del "%~dp0%installer%"

goto PipCheck
:PipCheck

python -m ensurepip --default-pip
python -m pip --version 2>nul
if errorlevel 1 (
  echo PIP not found, please ensure it's installed.
  pause
  exit /b 1
) else (
  goto GetDriveLetter
)

:GetDriveLetter
echo. 
:tryAgain
set /p drive=Please enter the path where you want to install ARG: 
set installPath=%drive%\ARG

call :validateFolder %installPath% installPath
rem echo %installPath%

set /p ays="ARG will be instaled in the "%installPath%" folder, is this correct (Y or [N])?"
IF /i "%ays%" == "Y"  GOTO CONTINUE
GOTO tryAgain
:CONTINUE

if not exist "%installPath%" (
    mkdir "%installPath%"
    cd "%installPath%"
    echo ARG folder created at %installPath%
) else (
    cd "%installPath%"
    echo ARG folder already exists at %installPath%. Proceeding with installation.
)

:: Copy Files
echo. 
echo copying essential files ... 
if not exist "%installPath%\template" (
    mkdir "%installPath%\template"
)

@echo off
echo @@echo off> arg.bat
echo set config=%%1>> arg.bat
echo.>> arg.bat
echo call .\.venv\scripts\activate>> arg.bat
echo .\.venv\scripts\python.exe %installPath%/src/cast_arg/main.py -c %%config%%>> arg.bat

echo arg.bat file created successfully.


:: Copying essential files to the destination folder
copy "%~dp0arg.bat" "%installPath%"
copy "%~dp0cause.json" "%installPath%"
copy "%~dp0text_replace.json" "%installPath%"
copy "%~dp0*config.json" "%installPath%"
copy "%~dp0README.md" "%installPath%"
copy "%~dp0*.pptx" "%installPath%\template"
copy "%~dp0*.xlsx" "%installPath%\template"
copy "%~dp0requirements.txt" "%installPath%"
robocopy "%~dp0src\cast_arg" "%installPath%\src\cast_arg" /E

:VenvSetup
echo. 
echo creating virtual environment ... 
echo. 
set CODE_FOLDER=%installPath%
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

:: Activate the virtual environment
call "%CODE_FOLDER%\.venv\Scripts\activate"

pip install -r requirements.txt

goto End

:VenvFail
echo Unable to install virtual environment
goto Usage

:Usage
echo Usage: install.bat
pause
exit /b 0

:End
echo ARG has been successfully installed at %CODE_FOLDER%
pause
exit /b 0


:validateFolder
set "var=%~1"

:: Determine absolute or relative
echo(!var:^"=!|findstr /i "^[A-Z]:[\\] ^[\\][\\]" >nul && set "type=absolute" || set "type=relative"

:: Determine file or folder or not exists
for /f eol^=^ delims^= %%F in ("!var!") do (
  for /f "tokens=1,2 delims=d" %%A in ("-%%~aF") do if "%%B" neq "" (
	set t=folder
  ) else if "%%A" neq "-" (
	set t=file
  ) else (
	set t=folder
  )
)
if "%type%"=="relative" (
	if "%t%"=="file" (
		echo input must be a folder
		goto error
	) else (
		if "%var:~1,1%" == ":" (
			set var=%var:~0,2%\%var:~2%
		) else (
			set var=%cd%\%var%
		)
	)
)

rem echo %var%
set "%~2=%var%" 
exit /b 0