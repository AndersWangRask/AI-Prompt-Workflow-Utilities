<#
.SYNOPSIS
    Removes the context menu entries created by setup.ps1 and deletes the symbolic link.

.DESCRIPTION
    This script removes the context menu entries for .docx and .pdf files created by setup.ps1.
    It also deletes the symbolic link created by setup.ps1 if it exists.

.NOTES
    - The script requires elevated privileges (run as administrator).

.EXAMPLE
    .\remove.ps1
#>

# Check for elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Script is not running with elevated privileges. Restarting with elevation..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Define the registry paths
$ContextMenuPathDocx = "HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\CreateMarkdownAIPrompt"
$ContextMenuPathPdf = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\CreateMarkdownAIPrompt"

# Remove context menu entry for .docx files
if (Test-Path -Path $ContextMenuPathDocx) {
    Remove-Item -Path $ContextMenuPathDocx -Recurse -Force
    Write-Host "Removed context menu entry for .docx files." -ForegroundColor Green
} else {
    Write-Host "Context menu entry for .docx files does not exist." -ForegroundColor Yellow
}

# Remove context menu entry for .pdf files
if (Test-Path -Path $ContextMenuPathPdf) {
    Remove-Item -Path $ContextMenuPathPdf -Recurse -Force
    Write-Host "Removed context menu entry for .pdf files." -ForegroundColor Green
} else {
    Write-Host "Context menu entry for .pdf files does not exist." -ForegroundColor Yellow
}

# Determine the Program Files directory
$ProgramFilesDir = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)
$TargetDir = "$ProgramFilesDir\AI-Prompt-Workflow-Utilities"

# Remove the junction if it exists
if (Test-Path -Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force
    Write-Host "Removed symbolic link: $TargetDir" -ForegroundColor Green
} else {
    Write-Host "Symbolic link does not exist: $TargetDir" -ForegroundColor Yellow
}

Write-Host "Removal script completed." -ForegroundColor Cyan
