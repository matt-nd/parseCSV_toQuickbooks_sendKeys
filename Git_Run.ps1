# Set the local repository path and GitHub repo URL
$repoPath = "C:\Users\matt\Scripts\parseCSV_toQuickbooks_sendKeys"
$repoURL = "https://github.com/matt-nd/parseCSV_toQuickbooks_sendKeys.git"
$scriptName = "SafeRun.ps1"

# Ensure Git is installed
if (-Not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Host "Error: Git is not installed. Please install Git." -ForegroundColor Red
    exit 1
}

# Step 1: Check if $repoPath exists and is a valid directory
if (-Not (Test-Path $repoPath -PathType Container)) {
    Write-Host "Repository folder does not exist. Creating and cloning repo..."
    git clone $repoURL $repoPath
} else {
    Write-Host "Repository folder exists. Checking if it's a valid Git repository..."
    
    # Step 2: Check if $repoURL is accessible
    try {
        $response = Invoke-WebRequest -Uri $repoURL -Method Head -ErrorAction Stop
        if ($response.StatusCode -ne 200) {
            Write-Host "Error: Unable to reach repository URL." -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "Error: GitHub repository URL is unreachable." -ForegroundColor Red
        exit 1
    }

    # Step 3: Check if it's a valid Git repository
    Set-Location $repoPath
    $gitCheck = git rev-parse --is-inside-work-tree 2>$null
    if ($gitCheck -ne "true") {
        Write-Host "Invalid Git repository. Deleting and recloning..."
        Remove-Item -Recurse -Force $repoPath
        git clone $repoURL $repoPath
    } else {
        # Step 4: Update the repository if needed
        Write-Host "Valid Git repository found. Checking for updates..."
        git fetch origin
        $localCommit = git rev-parse HEAD
        $remoteCommit = git rev-parse origin/main

        if ($localCommit -ne $remoteCommit) {
            Write-Host "Updates found. Pulling latest changes..."
            git reset --hard origin/main
            git pull
        } else {
            Write-Host "No updates needed. Repository is up-to-date."
        }
    }
}

# Run the PowerShell script
$scriptPath = "$repoPath\$scriptName"
if (Test-Path $scriptPath) {
    Write-Host "Running $scriptName..."
    Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$scriptPath`""
} else {
    Write-Host "Error: Script '$scriptName' not found!" -ForegroundColor Red
    exit 1
}
