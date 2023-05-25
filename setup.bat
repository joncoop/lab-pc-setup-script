@echo off
setlocal enabledelayedexpansion


rem URLs of the installer files
set "installers[0]=http://colorcop.net/tools/colorcop/colorcop-setup.exe"
set "installers[1]=https://www.bluej.org/download/files/BlueJ-windows-510.msi"
set "installers[2]=https://www.python.org/ftp/python/3.11.3/python-3.11.3-amd64.exe"
set "installers[3]=https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"

rem Application installer flags
set "flags[0]=/silent /MERGETASKS=desktopicon"
set "flags[1]=/qn"
set "flags[2]=/quiet InstallAllUsers=1 PrependPath=1 Include_test=0"
set "flags[3]=--silent --desktop_shortcut=false --gsuite_shortcuts=false"


rem User profiles to keep (must be separated by commas)
set "profilesToKeep=Administrator,Public,setup,Temp"

rem Python packages to install
set "packages=beautifulsoup4 matplotlib nympy pygame pyinstaller python-docx requests"

REM List of desktop icons to keep
set "keepIcons=HelpDesk.exe" "Audacity.lnk" "BlueJ.lnk" "Chrome.lnk" "Color Cop.lnk" "HelpDesk.exe" "Idle.lnk" "Microsoft Edge.lnk" "Visual Studio Code.lnk"

set "doneAction=restart"  REM Possible values: restart, shutdown. No value will keep the user logged in.


rem ******* END OF CONFIGURATION. DO NOT MODIFY BELOW THIS LINE! *******

echo Downloading installer files...

set "scriptDir=%~dp0"
set "installerDir=%scriptDir%installers"

mkdir "%installerDir%"

for /L %%i in (0,1,3) do (
    set "installers[%%i]=!installers[%%i]!"
    for %%F in ("!installers[%%i]!") do (
        set "filename=%%~nxF"
        echo Downloading installer %%i: !filename!
        bitsadmin.exe /transfer "InstallerDownload_%%i" "!installers[%%i]!" "%installerDir%\!filename!"
        echo Installer %%i downloaded successfully.
        set "installerPaths[%%i]=!filename!"
    )
)

echo Installing applications...

for /L %%i in (0,1,3) do (
    set "installerPath=%installerDir%\!installerPaths[%%i]!"
    set "flags[%%i]=!flags[%%i]!"

    if "!installerPath:~-4!"==".msi" (
        echo Installing MSI installer: !installerPath!
        msiexec.exe /i "!installerPath!" !flags[%%i]!
    ) else (
        for /F "tokens=1,*" %%a in ('echo !installerPath!') do (
            echo Installing %%a with flags !flags[%%i]!
            start "" /wait "%%a" !flags[%%i]! %%b
        )
    )
)

rem Create a Desktop shortcut for IDLE because the installer doesn't provide that option
copy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Python 3.11\IDLE (Python 3.11 64-bit).lnk" "C:\Users\Public\Desktop"
ren "C:\Users\Public\Desktop\IDLE (Python 3.11 64-bit).lnk" "IDLE.lnk"
echo IDLE shortcut added to Desktop.

rem Delete the installer directory
echo Deleting installer directory...
rd /s /q "%installerDir%"


rem Install python packages that may be used in class

echo Installing Python packages...

rem Update the PATH variable to include the location of pip
setx PATH "%PATH%;C:\Program Files\Python311\Scripts" /M

rem Wait for a moment to allow the PATH update to propagate
timeout /t 2 >nul

for %%P in (%packages%) do (
    echo Installing %%P...
    pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org %%P
    if not errorlevel 1 (
        echo Installation of %%P completed successfully.
    ) else (
        echo Failed to install %%P.
    )
)


rem Add shortcuts to Desktop for applications useful in computer programming, remove others

REM Unhide everything so I can see if this script misses anything hidden by previous hiding scripts
attrib -h -s "C:\Users\Public\Desktop\*" /A:-"C:\Users\Public\Desktop\desktop.ini"

REM Get the list of all desktop icons in the User/Public folder
for /f "tokens=*" %%i in ('dir "C:\Users\Public\Desktop\*" /b') do (
    REM Check if the current icon is in the list of icons to keep
    echo %keepIcons% | findstr /i "\<%%i\>" > nul
    if not errorlevel 1 (
        REM Icon is in the list, do not hide it
        attrib -s -h "C:\Users\Public\Desktop\%%i"
    ) else (
        REM Icon is not in the list, hide it
        attrib +s +h "C:\Users\Public\Desktop\%%i"
    )
)


rem Free space by deleting student profiles that have not been used within the past year

echo Deleting old student profiles...

for /F "delims=" %%G in ('dir "C:\Users" /B /AD') do (
    rem Check if profile is in keep list
    set "skipProfile="
    for %%P in (%profilesToKeep%) do (  
        if /I "%%G"=="%%P" set "skipProfile=1"
    )
    
    rem Check if the profile has been modified within the past year
    forfiles /P "C:\Users\%%G" /M * /D -365 /C "cmd /c exit 1" >nul 2>&1
    if not errorlevel 1 set "skipProfile=1"

    rem Delete profiles
    if not defined skipProfile (
        echo Deleting profile: %%G
        rd /s /q "C:\Users\%%G"
    ) else (
        echo Skipping profile: %%G
    )
)

endlocal

REM Perform the specified action
if "%doneAction%"=="restart" (
    echo Restarting in 5 seconds...
    shutdown /r /t 5
) else if "%doneAction%"=="shutdown" (
    echo Shutting down in 5 seconds...
    shutdown /s /t 5
) else (
    echo Invalid action specified. No further action taken.
    pause
)
