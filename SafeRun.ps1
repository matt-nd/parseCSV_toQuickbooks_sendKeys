# Get the directory where the script is running
$scriptDir = $PSScriptRoot

# Start the Kill Switch GUI
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$scriptDir\KillSwitch.ps1`"" 

# Start QuickBooks automation with proper logging
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"$scriptDir\PasteInvoices.ps1`"" -WindowStyle Normal
