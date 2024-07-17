<#
.SYNOPSIS
    Downloads and installs the latest version of Pandoc if not already installed or outdated, verifies the installation, and adds pandoc.exe to the system PATH if necessary.

.DESCRIPTION
    This script queries the GitHub API to find the latest release of Pandoc, checks the installed version (if any),
    downloads and installs Pandoc if necessary, verifies that pandoc.exe is installed in the expected location,
    and adds pandoc.exe to the system PATH if it is not already there.

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

# Function to get the installed Pandoc version
function Get-InstalledPandocVersion {
    try {
        $pandocVersionOutput = & pandoc --version 2>$null
        if ($pandocVersionOutput) {
            $pandocVersion = $pandocVersionOutput | Select-String -Pattern "pandoc.exe\s(\d+\.\d+\.\d+\.\d+|\d+\.\d+\.\d+)" | ForEach-Object { $_.Matches.Groups[1].Value }
            return $pandocVersion
        }
    } catch {
        return $null
    }
    return $null
}

# Function to remove any Pandoc paths from the user profile directory
function Remove-UserPandocPath {
    $userPandocPath = "C:\Users\$env:USERNAME\AppData\Local\Pandoc"
    $envPath = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::User)
    $updatedPath = ($envPath -split ';') -ne $userPandocPath
    $newPath = ($updatedPath -join ';')
    [System.Environment]::SetEnvironmentVariable("Path", $newPath, [System.EnvironmentVariableTarget]::User)
    Write-Host "Removed user Pandoc path: $userPandocPath" -ForegroundColor Yellow
}

# Function to uninstall Pandoc if it is installed
function Uninstall-Pandoc {
    Write-Host "Uninstalling existing Pandoc installation..." -ForegroundColor Yellow
    Start-Process msiexec.exe -ArgumentList "/x pandoc /quiet /norestart" -Wait
    Write-Host "Uninstallation complete." -ForegroundColor Yellow
}

# Remove any existing Pandoc paths in the user profile directory
Remove-UserPandocPath

# Get the latest release information
Write-Host "Fetching the latest Pandoc release information..." -ForegroundColor Cyan
$latestRelease = Get-LatestPandocRelease
$latestVersion = $latestRelease.tag_name.TrimStart('v')
$installerAsset = $latestRelease.assets | Where-Object { $_.name -like "*windows-x86_64.msi" }

if (-not $installerAsset) {
    Write-Host "Could not find a suitable installer for the current platform." -ForegroundColor Red
    exit
}

$pandocUrl = $installerAsset.browser_download_url
$installerPath = "$env:TEMP\pandoc-installer.msi"

# Output the temporary location where the MSI is downloaded
Write-Host "Downloading Pandoc installer to $installerPath" -ForegroundColor Cyan
Invoke-WebRequest -Uri $pandocUrl -OutFile $installerPath

# Check the installed Pandoc version
$installedVersion = Get-InstalledPandocVersion

if ($installedVersion) {
    Write-Host "Installed Pandoc version: $installedVersion" -ForegroundColor Yellow
} else {
    Write-Host "Pandoc is not currently installed." -ForegroundColor Yellow
}

if ($installedVersion -eq $latestVersion) {
    Write-Host "Pandoc is already up to date (version $installedVersion)." -ForegroundColor Green
} else {
    # Uninstall existing Pandoc if installed
    if ($installedVersion) {
        Uninstall-Pandoc
    }
    
    # Launch the installer and specify the installation directory for all users
    Write-Host "Launching Pandoc installer..." -ForegroundColor Cyan
    Start-Process msiexec.exe -ArgumentList "/i `"$installerPath`" /quiet /norestart ALLUSERS=1" -Wait

    # Verify the installation
    $pandocPath = "C:\Program Files\Pandoc\pandoc.exe"
    if (Test-Path -Path $pandocPath) {
        Write-Host "Pandoc installed successfully." -ForegroundColor Green
    } else {
        Write-Host "Pandoc installation failed." -ForegroundColor Red
        exit
    }
}

# Check the updated Pandoc version
$installedVersion = Get-InstalledPandocVersion
Write-Host "Updated Pandoc version: $installedVersion" -ForegroundColor Yellow

# Add Pandoc to the system PATH if not already present or if the PATH was not updated
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
