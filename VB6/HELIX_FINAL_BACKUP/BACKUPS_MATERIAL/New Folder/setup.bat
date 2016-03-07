@echo off
cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo This FCN updates the ViewCast Niagara SCX. 
echo .
echo Please follow these steps to update the software:
echo  1) Stop the Niagara service
echo  2) Run this installer
echo  3) Restart the Niagara service
echo .
echo Press the CTRL-C keys to abort the intall now or
pause

if EXIST "%windir%\System32\ViewCast\SCX\FCN04\Backup\RealProducer.dll" goto NoBackup
cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo Backing the current ViewCast files.
echo .
echo Creating backup directory:
echo     %windir%\System32\ViewCast\SCX
echo .
mkdir "%windir%\System32\ViewCast"
mkdir "%windir%\System32\ViewCast\SCX"
mkdir "%windir%\System32\ViewCast\SCX\FCN04"
mkdir "%windir%\System32\ViewCast\SCX\FCN04\Backup"
echo .
echo Backing up files:
echo     From:  %windir%\System32
echo     To:    %windir%\System32\ViewCast\SCX\FCN04\Backup
echo .
copy  "%windir%\System32\RealProducer.dll"	"%windir%\System32\ViewCast\SCX\FCN04\Backup"
echo .
echo Please verify that the three files were copied without errors.
echo .
echo Press the CTRL-C keys to abort the intall now or
pause
:NoBackup

cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo Copying updated files to %windir%\System32
copy RealProducer.dll        "%windir%\System32"
copy Readme_SCX_FCN04.txt   "%windir%\System32\ViewCast\SCX\FCN04"
echo .
echo Please verify that the four files were copied without
echo errors.
echo .
echo Press the CTRL-C keys to abort the intall now or
pause

cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo The setup program will now register the files.
echo After you press a key, you will see two
echo windows that indicates the success of the operations.
echo .
pause

rem ReRegister the Dll

cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo Please verify that the operations were successful.
%windir%\system32\Regsvr32 "%windir%\system32\RealProducer.dll"
echo .
echo Press the CTRL-C keys to abort the intall now or
pause

cls
echo -------------------------------------------------------
echo SCX 3.2.63.1.FCN04                           04/01/2004
echo ViewCast Corporation
echo -------------------------------------------------------
echo .
echo Installation of the Niagara SCX FCN 04 is complete.
echo .
pause
exit
rem - end
