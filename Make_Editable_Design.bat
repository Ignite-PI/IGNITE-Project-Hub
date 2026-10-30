@echo off
rem Drag a newsletter picture (PNG/JPG from Canva) onto this file to make its text editable.
if "%~1"=="" (
  echo Drag your newsletter picture ^(PNG or JPG^) onto this file.
  pause
  exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_editable_design.ps1" -Image "%~1"
pause
