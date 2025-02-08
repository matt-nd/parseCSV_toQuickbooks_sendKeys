# Set the local repository path and GitHub repo URL
$repoPath = "C:\Users\matt\Scripts\parseCSV_toQuickbooks_sendKeys"
$repoURL = "https://github.com/matt-nd/parseCSV_toQuickbooks_sendKeys.git"
$scriptName = "SafeRun.ps1"

# Function to check if Git is installed
function Check-GitInstalled {
    return (Get-Command git -ErrorAction SilentlyContinue) -ne $null
}

# Function to install Git if not found
function Install-Git {
    Write-Host "Git is not installed. Installing now..." -ForegroundColor Yellow
    $gitInstaller = "$env:TEMP\Git-Installer.exe"
    $gitDownloadURL = "https://github.com/git-for-windows/git/releases/latest/download/Git-64-bit.exe"

    # Download Git Installer
    Invoke-WebRequest -Uri $gitDownloadURL -OutFile $gitInstaller
    Write-Host "Git downloaded successfully."

    # Run the installer silently
    Start-Process -FilePath $gitInstaller -ArgumentList "/VERYSILENT /NORESTART /SUPPRESSMSGBOXES" -Wait
    Write-Host "Git installed successfully. Adding Git to PATH..."

    # Refresh PowerShell session to recognize Git
    $env:Path += ";C:\Program Files\Git\bin;C:\Program Files\Git\cmd"
}

# Step 1: Check if Git is installed, install if missing
if (-Not (Check-GitInstalled)) {
    Install-Git

    # Verify Git installation after installing
    if (-Not (Check-GitInstalled)) {
        Write-Host "Error: Git installation failed. Please install manually from https://git-scm.com/downloads" -ForegroundColor Red
        exit 1
    }
}