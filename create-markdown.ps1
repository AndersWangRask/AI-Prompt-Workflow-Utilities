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
if ($Extension -ne ".docx" -and $Extension -ne ".pdf") {
    Write-Host "Invalid file type. Only .docx and .pdf files are supported." -ForegroundColor Red
    exit
}

# Get directory and file name without extension
$Directory = [System.IO.Path]::GetDirectoryName($FilePath)
$FileNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
$TargetFilePath = "$Directory\$FileNameWithoutExtension.md"

# Check for archive script in the same directory
$ArchiveScriptPath = Get-ChildItem -Path $Directory -Filter "archive.ps1" -Recurse -ErrorAction SilentlyContinue

if ($ArchiveScriptPath) {
    Write-Host "Executing archive script: $ArchiveScriptPath" -ForegroundColor Yellow
    & $ArchiveScriptPath.FullName
} else {
    Write-Host "Archive script not found in the directory." -ForegroundColor Yellow
}

# Convert the file to Markdown using Pandoc
Write-Host "Converting $FilePath to $TargetFilePath" -ForegroundColor Green
if ($Extension -eq ".docx") {
    pandoc "$FilePath" -f docx -t markdown -s -o "$TargetFilePath"
} elseif ($Extension -eq ".pdf") {
    pandoc "$FilePath" -f pdf -t markdown -s -o "$TargetFilePath"
}

if ($?) {
    Write-Host "Conversion successful: $TargetFilePath" -ForegroundColor Green
} else {
    Write-Host "Conversion failed." -ForegroundColor Red
}
