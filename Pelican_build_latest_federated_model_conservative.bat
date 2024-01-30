@echo off
TITLE Pelican FM Automation - Build Latest Federated Models (Conservative)
REM setlocal ENABLEEXTENSIONS
REM setlocal ENABLEDELAYEDEXPANSION

set scriptDir=%~dp0
set maindir=%cd%

cd "%scriptDir%\Scripts"
powershell -noprofile -ExecutionPolicy Bypass -File Pelican_main_conservative.ps1
cd %maindir%

pause