@echo off
set scriptDir=%~dp0
cd %scriptDir%
powershell -noprofile -ExecutionPolicy Bypass -File Pelican-Basic-FM.ps1
pause