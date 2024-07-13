# Get the current directory where the script is located
$CurrentDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define the archive directory name
$ArchiveDir = "$CurrentDir\Archive"

# Define the log file name
$LogFile = "$CurrentDir\archive_log.txt"

# Function to log messages
Function Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
    $Message | Out-File -FilePath $LogFile -Append
}

# Check if the Archive directory exists, if not, create it
if (-not (Test-Path -Path $ArchiveDir -PathType Container)) {
    New-Item -ItemType Directory -Path $ArchiveDir
    Log "Created Archive directory: $ArchiveDir" "Green"
}

# Get the full path of the script file
$ScriptFullPath = $MyInvocation.MyCommand.Definition

# Get the list of files in the current directory, excluding the script file, sub-directories, and the log file
$Files = Get-ChildItem -Path $CurrentDir -File | Where-Object { $_.FullName -ne $ScriptFullPath -and $_.FullName -ne $LogFile }

# Loop through each file and copy it to the Archive directory with the custom date-time format
foreach ($File in $Files) {
    $LastModified = $File.LastWriteTime.ToString("yyyy-MM-dd-HHmm")
    $ArchiveFileName = "$LastModified-$($File.Name)"
    $ArchiveFilePath = "$ArchiveDir\$ArchiveFileName"
    
    # Check if the file already exists in the Archive directory
    if (-not (Test-Path -Path $ArchiveFilePath)) {
        Copy-Item -Path $File.FullName -Destination $ArchiveFilePath
        Log "Copied $($File.Name) to $ArchiveFileName" "Green"
    } else {
        Log "Skipped $($File.Name) - already exists as $ArchiveFileName" "Yellow"
    }
}

# Pause at the end to see the output
Log "Press any key to exit..." "Cyan"
[System.Console]::ReadKey() | Out-Null
