<#
.SYNOPSIS
    Converts .pdf or .md files to DOCX format (.docx) using Pandoc.

.DESCRIPTION
    This script takes a file path as a parameter and converts the file to DOCX format (.docx).
    It supports .pdf and .md files. If an archive script (archive.ps1) is found in the same directory,
    it executes the archive script before performing the conversion.

.PARAMETER FilePath
    The full path of the .pdf or .md file to be converted.

.NOTES
    - The script requires Pandoc to be installed and available in the system PATH.
    - Pandoc can be downloaded from: https://pandoc.org/installing.html

.EXAMPLE
    .\ConvertToDocx.ps1 -FilePath "C:\path\to\your\file.pdf"

#>

param (
    [string]$FilePath
)

# Check if the file exists
if (-not (Test-Path -Path $FilePath)) {
    Write-Host "File does not exist: $FilePath" -ForegroundColor Red
    exit
}

# Get file extension
$Extension = [System.IO.Path]::GetExtension($FilePath).ToLower()

# Validate file extension
if ($Extension -ne ".pdf" -and $Extension -ne ".md") {
    Write-Host "Invalid file type. Only .pdf and .md files are supported." -ForegroundColor Red
    exit
}

# Get directory and file name without extension
$Directory = [System.IO.Path]::GetDirectoryName($FilePath)
$FileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$TargetFilePath = "$Directory\$FileNameWithoutExtension.docx"

# Check for archive script in the same directory
$ArchiveScriptPath = Get-ChildItem -Path $Directory -Filter "archive.ps1" -Recurse -ErrorAction SilentlyContinue

if ($ArchiveScriptPath) {
    Write-Host "Executing archive script: $ArchiveScriptPath" -ForegroundColor Yellow
    & $ArchiveScriptPath.FullName
} else {
    Write-Host "Archive script not found in the directory." -ForegroundColor Yellow
}

# Convert the file to DOCX using Pandoc
Write-Host "Converting $FilePath to $TargetFilePath" -ForegroundColor Green
if ($Extension -eq ".pdf") {
    pandoc "$FilePath" -f pdf -t docx -s -o "$TargetFilePath"
} elseif ($Extension -eq ".md") {
    pandoc "$FilePath" -f markdown -t docx -s -o "$TargetFilePath"
}

if ($?) {
    Write-Host "Conversion successful: $TargetFilePath" -ForegroundColor Green
} else {
    Write-Host "Conversion failed." -ForegroundColor Red
}
