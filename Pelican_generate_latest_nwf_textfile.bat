@echo off
TITLE Pelican FM Automation - Generate Latest Federated Model Text Files
REM setlocal ENABLEEXTENSIONS
REM setlocal ENABLEDELAYEDEXPANSION

set scriptDir=%~dp0
set maindir=%cd%

cd "%scriptDir%\Scripts"
powershell -noprofile -ExecutionPolicy Bypass -File Pelican_generate_nwf_textfile.ps1
cd %maindir%

pause