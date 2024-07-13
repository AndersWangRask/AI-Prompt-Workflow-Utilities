<#
.SYNOPSIS
    Creates a symbolic link to the directory where this script is executed from and adds a context menu entry for .docx and .pdf files to create a Markdown file using create-markdown.ps1.

.DESCRIPTION
    This script creates a symbolic link from the directory where it is executed to a specified target directory.
    It also adds a context menu entry for .docx and .pdf files to create a Markdown file using create-markdown.ps1.

.NOTES
    - The script requires elevated privileges (run as administrator).
    - Ensure that the target directory does not already exist, or it will be replaced.

.EXAMPLE
    .\setup.ps1
#>

# Check for elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Script is not running with elevated privileges. Restarting with elevation..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Get the directory where the script is located
$SourceDir = $PSScriptRoot
$ProgramFilesDir = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles)
$TargetDir = "$ProgramFilesDir\AI-Prompt-Workflow-Utilities"

# Remove the target directory if it exists
if (Test-Path -Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force
}

# Create the symbolic link for the directory
New-Item -ItemType Junction -Path $TargetDir -Target $SourceDir

# Output the result
Write-Host "Symbolic link created from $SourceDir to $TargetDir" -ForegroundColor Green

# Define the script path for create-markdown.ps1
$CreateMarkdownScript = "$TargetDir\create-markdown.ps1"

# Define the registry paths
$ContextMenuPathDocx = "HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\CreateMarkdownAIPrompt"
$ContextMenuPathPdf = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\CreateMarkdownAIPrompt"

# Add context menu entry for .docx files
New-Item -Path $ContextMenuPathDocx -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathDocx -Name "(Default)" -Value "Create Markdown (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathDocx\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathDocx\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$CreateMarkdownScript`" `"%1`"" | Out-Null

# Add context menu entry for .pdf files
New-Item -Path $ContextMenuPathPdf -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathPdf -Name "(Default)" -Value "Create Markdown (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathPdf\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathPdf\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$CreateMarkdownScript`" `"%1`"" | Out-Null

# Output the result
Write-Host "Context menu entries created for .docx and .pdf files." -ForegroundColor Green
