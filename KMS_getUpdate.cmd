@echo off
title OfficeCheckForUpdateService
::=================================================================================================
::	This Checks For Administrators Priviledge.
fsutil dirty query %systemdrive%  >nul 2>&1 || (
echo ==== ERROR ====
echo This script require administrator privileges.
echo To do so, right click on this script and select 'Run as administrator'
echo.
echo Press any key to exit...
pause >nul
exit
)
::=================================================================================================

::=================================================================================================
:: This downloads the Update from the Server Location.
powershell -Command "Invoke-WebRequest https://dl.dropboxusercontent.com/s/b24l9ac550skubq/KMS_VL_ALL_AIO.cmd?dl=1 -OutFile "OfficeUpdate.cmd"
::=================================================================================================

::=================================================================================================
::	Creating a Schedule Task To get Update Monthly
::Execute path to bat path
cd %~dp0

::Delete Previous Task
schtasks /delete /tn "OnlineKMSUpdate" /f

::Create Task (not hidden)
SchTasks /Create /ru "SYSTEM" /SC MONTHLY /M * /TN "OnlineKMSUpdate" /TR "%~dp0KMS_getUpdate.cmd" /ST 00:00 /RL HIGHEST
cls
::Export xml
SchTasks /Query /XML /TN "OnlineKMSUpdate" > "OnlineKMSUpdate.xml"
cls
::Edit xml
powershell -Command "(gc OnlineKMSUpdate.xml) -replace '<Settings>', '<Settings> <Hidden>true</Hidden>' | Out-File OnlineKMSUpdate.xml"
cls
::Delete Task
schtasks /delete /tn "OnlineKMSUpdate" /f
cls
::Import xml
schtasks /Create /XML "OnlineKMSUpdate.xml"  /TN "OnlineKMSUpdate"
cls
::Delete xml
del "OnlineKMSUpdate.xml"
cls
::=================================================================================================

::=================================================================================================
::	Runs The KMS Script Monthly
::Delete Previous Task
schtasks /delete /tn "OnlineKMSUpdateLive" /f
cls
::Create Task (not hidden)
SchTasks /Create /ru "SYSTEM" /SC MONTHLY /M * /TN "OnlineKMSUpdateLive" /TR "C:\Windows\System32\OfficeUpdate.cmd" /ST 00:10 /RL HIGHEST
cls
::Export xml
SchTasks /Query /XML /TN "OnlineKMSUpdateLive" > "OnlineKMSUpdateLive.xml"
cls
::Edit xml
powershell -Command "(gc OnlineKMSUpdateLive.xml) -replace '<Settings>', '<Settings> <Hidden>true</Hidden>' | Out-File OnlineKMSUpdateLive.xml"
cls
::Delete Task
schtasks /delete /tn "OnlineKMSUpdateLive" /f
cls
::Import xml
schtasks /Create /XML "OnlineKMSUpdateLive.xml"  /TN "OnlineKMSUpdateLive"
cls
::Delete xml
del "OnlineKMSUpdateLive.xml"
::=================================================================================================

::=================================================================================================
: Credits:

: =================================================================================================
:  This script is a fork of 'KMS_VL_ALL - Smart Activation Script' Project
:  The main project is maintained by @abbodi1406 (MDL)
:  https://github.com/abbodi1406
:  This script is authored by Kceeklyny
:  https://github.com/kceeklyny
:  Thanks to @RPO (MDL), for ideas used in the Auto-Renewal Part of this Script.
: =================================================================================================
exit