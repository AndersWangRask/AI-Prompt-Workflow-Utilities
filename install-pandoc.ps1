<#
.SYNOPSIS
    Downloads and installs the latest version of Pandoc, verifies the installation, and adds pandoc.exe to the system PATH.

.DESCRIPTION
    This script queries the GitHub API to find the latest release of Pandoc, downloads the installer, launches the installer,
    waits for the installation to complete, verifies that pandoc.exe is installed in the expected location, and adds pandoc.exe
    to the system PATH if necessary.

.NOTES
    - The script requires elevated privileges (run as administrator).

.EXAMPLE
    .\install-pandoc.ps1
#>

# Check for elevation
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Script is not running with elevated privileges. Restarting with elevation..."
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# Function to get the latest release information from GitHub
function Get-LatestPandocRelease {
    $apiUrl = "https://api.github.com/repos/jgm/pandoc/releases/latest"
    $response = Invoke-RestMethod -Uri $apiUrl -Headers @{ "User-Agent" = "PowerShell" }
    return $response
}

# Get the latest release information
Write-Host "Fetching the latest Pandoc release information..." -ForegroundColor Cyan
$latestRelease = Get-LatestPandocRelease
$pandocVersion = $latestRelease.tag_name
$installerAsset = $latestRelease.assets | Where-Object { $_.name -like "*windows-x86_64.msi" }

if (-not $installerAsset) {
    Write-Host "Could not find a suitable installer for the current platform." -ForegroundColor Red
    exit
}

$pandocUrl = $installerAsset.browser_download_url
$installerPath = "$env:TEMP\pandoc-installer.msi"

# Download the Pandoc installer
Write-Host "Downloading Pandoc installer ($pandocVersion)..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $pandocUrl -OutFile $installerPath

# Launch the installer
Write-Host "Launching Pandoc installer..." -ForegroundColor Cyan
Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait

# Verify the installation
$pandocPath = "C:\Program Files\Pandoc\pandoc.exe"
if (Test-Path -Path $pandocPath) {
    Write-Host "Pandoc installed successfully." -ForegroundColor Green
} else {
    Write-Host "Pandoc installation failed." -ForegroundColor Red
    exit
}

# Add Pandoc to the system PATH if not already present
$envPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::Machine)
if (-not $envPath.Contains("C:\Program Files\Pandoc")) {
    Write-Host "Adding Pandoc to the system PATH..." -ForegroundColor Cyan
    [System.Environment]::SetEnvironmentVariable("Path", "$envPath;C:\Program Files\Pandoc", [System.EnvironmentVariableTarget]::Machine)
    Write-Host "Pandoc added to the system PATH. You may need to restart your terminal for the changes to take effect." -ForegroundColor Green
} else {
    Write-Host "Pandoc is already in the system PATH." -ForegroundColor Yellow
}

# Clean up the installer
Remove-Item -Path $installerPath -Force
Write-Host "Installation complete." -ForegroundColor Cyan
