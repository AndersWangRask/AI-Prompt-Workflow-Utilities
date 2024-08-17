<#
.SYNOPSIS
    Combines multiple Markdown files and converts the result to DOCX format.

.DESCRIPTION
    This script takes multiple Markdown (.md) file paths as input, combines them
    in alphabetical order into a single Markdown file, converts the combined file
    to DOCX format using the create-docx.ps1 script, and then deletes the
    combined Markdown file. The output DOCX file is placed in the same directory
    as the first input file.

.PARAMETER FilePaths
    An array of file paths to the Markdown files to be combined and converted.

.NOTES
    - Requires create-docx.ps1 to be in the same directory as this script.
    - Assumes Pandoc is installed and available in the system PATH.

.EXAMPLE
    .\combine-and-convert.ps1 -FilePaths "C:\path\to\file1.md", "C:\path\to\file2.md"
#>

param (
    [Parameter(Mandatory=$true)]
    [string[]]$FilePaths
)

# Function to write log messages
function Write-Log {
    param (
        [string]$Message
    )
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message"
}

# Check if all input files exist and are Markdown files
foreach ($file in $FilePaths) {
    if (-not (Test-Path $file)) {
        Write-Log "Error: File not found: $file"
        exit 1
    }
    if ([System.IO.Path]::GetExtension($file).ToLower() -ne ".md") {
        Write-Log "Error: File is not a Markdown file: $file"
        exit 1
    }
}

# Sort file paths alphabetically
$SortedFilePaths = $FilePaths | Sort-Object

# Generate output file name based on input files
$FirstFile = $SortedFilePaths[0]
$OutputDir = [System.IO.Path]::GetDirectoryName($FirstFile)
$OutputBaseName = if ($SortedFilePaths.Count -eq 1) {
    [System.IO.Path]::GetFileNameWithoutExtension($FirstFile)
} else {
    $FirstFileName = [System.IO.Path]::GetFileNameWithoutExtension($FirstFile)
    $LastFileName = [System.IO.Path]::GetFileNameWithoutExtension($SortedFilePaths[-1])
    "${FirstFileName}_to_${LastFileName}_combined"
}
$OutputDocxFile = Join-Path $OutputDir "${OutputBaseName}.docx"

# Create a temporary file for the combined Markdown content
$TempDir = [System.IO.Path]::GetTempPath()
$CombinedMarkdownFile = Join-Path $TempDir "combined_markdown_$(Get-Date -Format 'yyyyMMddHHmmss').md"

Write-Log "Combining Markdown files into: $CombinedMarkdownFile"

# Combine the content of all Markdown files
foreach ($file in $SortedFilePaths) {
    Get-Content $file | Add-Content $CombinedMarkdownFile
    "`n`n" | Add-Content $CombinedMarkdownFile  # Add two newlines between files
}

# Get the directory of the current script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path to the create-docx.ps1 script
$CreateDocxScript = Join-Path $ScriptDir "create-docx.ps1"

if (-not (Test-Path $CreateDocxScript)) {
    Write-Log "Error: create-docx.ps1 not found in the same directory as this script."
    exit 1
}

Write-Log "Converting combined Markdown to DOCX..."

# Call the create-docx.ps1 script to convert the combined Markdown to DOCX
& $CreateDocxScript -FilePath $CombinedMarkdownFile

# Check if the conversion was successful
if ($LASTEXITCODE -eq 0) {
    Write-Log "Conversion completed successfully."
    
    # Delete the temporary combined Markdown file
    Remove-Item $CombinedMarkdownFile
    Write-Log "Temporary combined Markdown file deleted."

    # Get the path of the created DOCX file (it will be in the same location as the combined Markdown file)
    $CreatedDocxFile = [System.IO.Path]::ChangeExtension($CombinedMarkdownFile, "docx")
    
    # Move and rename the created DOCX file to the desired output location
    Move-Item -Path $CreatedDocxFile -Destination $OutputDocxFile -Force
    Write-Log "Created DOCX file: $OutputDocxFile"
} else {
    Write-Log "Error: Conversion failed. Check the create-docx.ps1 script for more details."
}

Write-Log "Process completed."