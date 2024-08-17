@echo off
setlocal enabledelayedexpansion

rem Build the argument string
set args=
for %%I in (%*) do (
    if defined args (
        set "args=!args!, \"%%~fI\""
    ) else (
        set "args=\"%%~fI\""
    )
)

rem Call the PowerShell script with the constructed argument string
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%~dp0combine-and-convert.ps1' -FilePaths %args%"
rem pause