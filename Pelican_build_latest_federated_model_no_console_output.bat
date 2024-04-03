@echo off
TITLE Pelican FM Automation - Build Latest Federated Models
REM setlocal ENABLEEXTENSIONS
REM setlocal ENABLEDELAYEDEXPANSION

set scriptDir=%~dp0
set maindir=%cd%
set dateformat=%date:~10,4%%date:~7,2%%date:~4,2%

cd "%scriptDir%\Scripts"
powershell -noprofile -ExecutionPolicy Bypass -File Pelican_main.ps1 > Scripts\Logs\Console_output_%dateformat%.csv
cd %maindir%

pause