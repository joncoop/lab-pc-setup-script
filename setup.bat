@echo off
setlocal enabledelayedexpansion


rem Application installers and flags
rem URLs of the installer files
set "installers[0]=http://colorcop.net/tools/colorcop/colorcop-setup.exe"
set "installers[1]=https://www.python.org/ftp/python/3.11.3/python-3.11.3-amd64.exe"
set "installers[2]=https://www.bluej.org/download/files/BlueJ-windows-510.msi"
set "installers[3]=https://dl.google.com/drive-file-stream/GoogleDriveSetup.exe"

rem Application installer flags
set "flags[0]=/silent /MERGETASKS=desktopicon"
set "flags[1]=/quiet InstallAllUsers=1 PrependPath=1 Include_test=0"
set "flags[2]=/qn"
set "flags[3]=--silent --desktop_shortcut=false --gsuite_shortcuts=false"


rem User profiles to keep (must be separated by commas)
set "profilesToKeep=Administrator,Public,setup,Temp"


rem Python packages to install (separate with spaces)
set "packages=beautifulsoup4 matplotlib nympy pygame pyinstaller python-docx requests"


rem Desktop shortcuts to hide
set "shortcutsToHide="
set "shortcutsToHide=%shortcutsToHide% "Access.lnk""
set "shortcutsToHide=%shortcutsToHide% "Acrobat Reader.lnk""
set "shortcutsToHide=%shortcutsToHide% "Destiny Discover.url""
set "shortcutsToHide=%shortcutsToHide% "DisableScreensaver.appref-ms""
set "shortcutsToHide=%shortcutsToHide% "Dismissal.url""
set "shortcutsToHide=%shortcutsToHide% "DRC INSIGHT Online Assessments.lnk""
set "shortcutsToHide=%shortcutsToHide% "Excel.lnk""
set "shortcutsToHide=%shortcutsToHide% "GCSD Authentication.url""
set "shortcutsToHide=%shortcutsToHide% "Internet Explorer.lnk""
set "shortcutsToHide=%shortcutsToHide% "JL Mann Broadcast.lnk""
set "shortcutsToHide=%shortcutsToHide% "LockDown Browser Lab OEM.lnk""
set "shortcutsToHide=%shortcutsToHide% "Office.lnk""
set "shortcutsToHide=%shortcutsToHide% "OneNote.lnk""
set "shortcutsToHide=%shortcutsToHide% "NWEA Secure Testing Browser.lnk""
set "shortcutsToHide=%shortcutsToHide% "PowerPoint.lnk""
set "shortcutsToHide=%shortcutsToHide% "Publisher.lnk""
set "shortcutsToHide=%shortcutsToHide% "Raptor Staff SignIn.url""
set "shortcutsToHide=%shortcutsToHide% "S A S.url""
set "shortcutsToHide=%shortcutsToHide% "SCSecureBrowser.lnk""
set "shortcutsToHide=%shortcutsToHide% "Sub""
set "shortcutsToHide=%shortcutsToHide% "WIN Career Readiness.url""
set "shortcutsToHide=%shortcutsToHide% "Word.lnk""


rem ******* END OF CONFIGURATION. DO NOT MODIFY BELOW THIS LINE! *******

echo Downloading installer files...

set "scriptDir=%~dp0"
set "installerDir=%scriptDir%installers"

mkdir "%installerDir%"

for /L %%i in (0,1,3) do (
    set "installers[%%i]=!installers[%%i]!"
    for /F "delims=" %%F in ('echo !installers[%%i]!') do (
        set "url=%%F"
        set "extension=!url:~-3!"
        echo Downloading installer %%i with extension !extension!
        bitsadmin.exe /transfer "InstallerDownload_%%i" "!installers[%%i]!" "%installerDir%\Installer_%%i.!extension!"
        echo Installer %%i downloaded successfully.
    )
)

echo Installing applications...

for /L %%i in (0,1,3) do (
    set "installerPath=%installerDir%\Installer_%%i.!extension!"
    set "flags[%%i]=!flags[%%i]!"

    for /F "tokens=1,*" %%a in ('echo !installerPath!') do (
        echo Installing %%a with flags !flags[%%i]!

        if /I "!extension!"==".exe" (
            start "" /wait "%%a" !flags[%%i]! %%b
        ) else if /I "!extension!"==".msi" (
            start "" /wait msiexec.exe /i "%%a" !flags[%%i]! %%b
        ) else (
            echo Unsupported file type: !extension! Skipping installation of %%a.
        )
    )
)

rem Delete the installers once complete
rd /s /q "%installerDir%"


rem Create a Desktop shortcut for IDLE because the installer doesn't provide that option
copy "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Python 3.11\IDLE (Python 3.11 64-bit).lnk" "C:\Users\Public\Desktop"
ren "C:\Users\Public\Desktop\IDLE (Python 3.11 64-bit).lnk" "IDLE.lnk"
echo IDLE shortcut added to Desktop.


rem Install python packages that will be used in class

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


rem Add shortcuts to Desktop for applications useful in computer programming, remove others

REM Unhide everything so I can see if this script misses anything hidden by previous hiding scripts
attrib -h -s "C:\Users\Public\Desktop\*" /A:-"C:\Users\Public\Desktop\desktop.ini"

echo Hiding unneeded Desktop shortcuts...

for %%S in (%shortcutsToHide%) do (
    echo Hiding shortcut: %%~S
    attrib +h +s "C:\Users\Public\Desktop\%%~S"
)


endlocal

pause
