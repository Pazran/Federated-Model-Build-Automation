@echo off
TITLE Pelican FM Automation - Build Latest Federated Models
REM setlocal ENABLEEXTENSIONS
REM setlocal ENABLEDELAYEDEXPANSION

set scriptDir=%~dp0
set maindir=%cd%

cd "%scriptDir%\Scripts"
powershell -noprofile -ExecutionPolicy Bypass -File Pelican_main.ps1
cd %maindir%

pause