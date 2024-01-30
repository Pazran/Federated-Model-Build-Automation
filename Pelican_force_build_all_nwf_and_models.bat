@echo off
TITLE Pelican FM Automation - Force Build All Federated Models
REM setlocal ENABLEEXTENSIONS
REM setlocal ENABLEDELAYEDEXPANSION

set scriptDir=%~dp0
set maindir=%cd%

cd "%scriptDir%\Scripts"
powershell -noprofile -ExecutionPolicy Bypass -File Pelican_force_build_nwf_and_models.ps1
cd %maindir%

pause