<#
.SYNOPSIS
    Creates a symbolic link to the directory where this script is executed from and adds context menu entries for .docx and .pdf files to create a Markdown file using create-markdown.ps1.
    Also adds context menu entries for .md and .pdf files to create a DOCX file using ConvertToDocx.ps1.

.DESCRIPTION
    This script creates a symbolic link from the directory where it is executed to a specified target directory.
    It also adds context menu entries for .docx and .pdf files to create a Markdown file using create-markdown.ps1,
    and for .md and .pdf files to create a DOCX file using ConvertToDocx.ps1.

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
$TargetDir = "$ProgramFilesDir\\AI-Prompt-Workflow-Utilities"

# Remove the target directory if it exists
if (Test-Path -Path $TargetDir) {
    Remove-Item -Path $TargetDir -Recurse -Force
}

# Create the symbolic link for the directory
New-Item -ItemType Junction -Path $TargetDir -Target $SourceDir

# Output the result
Write-Host "Symbolic link created from $SourceDir to $TargetDir" -ForegroundColor Green

# Define the script paths
$CreateMarkdownScript = "$TargetDir\\create-markdown.ps1"
$ConvertToDocxScript = "$TargetDir\\create-docx.ps1"

# Define the registry paths for create-markdown.ps1
$ContextMenuPathDocxMarkdown = "HKCU:\\Software\\Classes\\SystemFileAssociations\\.docx\\shell\\CreateMarkdownAIPrompt"
$ContextMenuPathPdfMarkdown = "HKCU:\\Software\\Classes\\SystemFileAssociations\\.pdf\\shell\\CreateMarkdownAIPrompt"

# Add context menu entry for .docx files to create Markdown
New-Item -Path $ContextMenuPathDocxMarkdown -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathDocxMarkdown -Name "(Default)" -Value "Create Markdown (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathDocxMarkdown\\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathDocxMarkdown\\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$CreateMarkdownScript`" `"%1`"" | Out-Null

# Add context menu entry for .pdf files to create Markdown
New-Item -Path $ContextMenuPathPdfMarkdown -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathPdfMarkdown -Name "(Default)" -Value "Create Markdown (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathPdfMarkdown\\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathPdfMarkdown\\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$CreateMarkdownScript`" `"%1`"" | Out-Null

# Define the registry paths for ConvertToDocx.ps1
$ContextMenuPathMdDocx = "HKCU:\\Software\\Classes\\SystemFileAssociations\\.md\\shell\\CreateDocxAIPrompt"
$ContextMenuPathPdfDocx = "HKCU:\\Software\\Classes\\SystemFileAssociations\\.pdf\\shell\\CreateDocxAIPrompt"

# Add context menu entry for .md files to create DOCX
New-Item -Path $ContextMenuPathMdDocx -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathMdDocx -Name "(Default)" -Value "Create DOCX (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathMdDocx\\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathMdDocx\\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ConvertToDocxScript`" `"%1`"" | Out-Null

# Add context menu entry for .pdf files to create DOCX
New-Item -Path $ContextMenuPathPdfDocx -Force | Out-Null
Set-ItemProperty -Path $ContextMenuPathPdfDocx -Name "(Default)" -Value "Create DOCX (AI Prompt)" | Out-Null
New-Item -Path "$ContextMenuPathPdfDocx\\command" -Force | Out-Null
Set-ItemProperty -Path "$ContextMenuPathPdfDocx\\command" -Name "(Default)" -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$ConvertToDocxScript`" `"%1`"" | Out-Null

# Output the result
Write-Host "Context menu entries created for .docx, .pdf, and .md files." -ForegroundColor Green


# Setup for Combine and Convert to DOCX utility
# This section adds a shortcut to the SendTo menu for the combine-and-convert.bat script
# The utility allows users to select multiple Markdown files, combine them, and convert to a single DOCX file

# Define the path for the combine-and-convert.bat file
# This bat file acts as an intermediary between the SendTo menu and the PowerShell script
$CombineAndConvertBat = "$TargetDir\combine-and-convert.bat"

# Define the SendTo folder path
# The SendTo folder is a special Windows folder that integrates with the right-click context menu
$SendToFolder = [System.Environment]::GetFolderPath('ApplicationData') + "\Microsoft\Windows\SendTo"

# Create the shortcut in the SendTo folder
# This allows users to select multiple Markdown files, right-click, and choose this option from the SendTo menu
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$SendToFolder\Combine and Convert to DOCX.lnk")
$Shortcut.TargetPath = $CombineAndConvertBat
$Shortcut.Save()

# Provide feedback to the user about the creation of the SendTo shortcut
Write-Host "SendTo shortcut 'Combine and Convert to DOCX' created for combine-and-convert.bat" -ForegroundColor Green
Write-Host "You can now select multiple Markdown files, right-click, and use 'Send To' -> 'Combine and Convert to DOCX'" -ForegroundColor Green

# Clean up the COM object to free resources
# This is a good practice when working with COM objects in PowerShell
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WshShell) | Out-Null