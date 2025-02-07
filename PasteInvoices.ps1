# Dot-source the file to import its functions
. "C:\Users\matt\Desktop\InvoiceInput\Dialogs.ps1"   # Note the dot and space before the file path

# Check if the WinAPI type already exists before adding it
if (-not ([System.Management.Automation.PSTypeName]'WinAPI').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class WinAPI {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
   
}
"@
}

function Activate-Window {
    param ([IntPtr]$hWnd)

    $SW_RESTORE = 9 # Restore window if minimized
    [WinAPI]::ShowWindow($hWnd, $SW_RESTORE) | Out-Null
    [WinAPI]::SetForegroundWindow($hWnd) | Out-Null
}


function Get-RDPWindows {
    $script:windows = @()  # Declare as script-scoped to persist across callback executions
    
    $callback = {
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if ([WinAPI]::IsWindowVisible($hWnd)) {
            $sb = New-Object System.Text.StringBuilder(256)
            [void][WinAPI]::GetWindowText($hWnd, $sb, $sb.Capacity)

            $title = $sb.ToString()
            if ($title) {
                $script:windows += @{ Handle = $hWnd; Title = $title }
                Write-Host "Found Window: $title"
            }
        }

        return $true
    }

    $enumProc = [WinAPI+EnumWindowsProc]$callback
    [WinAPI]::EnumWindows($enumProc, [IntPtr]::Zero) | Out-Null

    if (0 -eq $windows.Count) {
        Write-Host "!! No windows found. Try running as Administrator." -ForegroundColor Red
    }

    return $script:windows
}


function Get-ClipboardContent {
    Add-Type -AssemblyName System.Windows.Forms
    return [System.Windows.Forms.Clipboard]::GetText()
}

function Convert-ClipboardData {
    param ($clipboardText, $selectedColumns)

    # Debugging: Print the clipboard content
    Write-Host "`n[DEBUG] Raw Clipboard Data:"
    Write-Host ($clipboardText | Out-String)

    # Split clipboard text into lines
    $rows = $clipboardText -split "`r?`n"
    $parsedData = @()

    foreach ($row in $rows) {
        if ($row -match "\S") {  # Ignore empty lines
            $columns = $row -split "`t"  # Split by tab characters
            $filteredRow = @()

            # Keep only the selected columns
            foreach ($colIndex in $selectedColumns) {
                if ($colIndex -ge 0 -and $colIndex -lt $columns.Length) {
                    $filteredRow += $columns[$colIndex]
                } else {
                    $filteredRow += ""  # Add an empty string if column is out of range
                }
            }
            
            $parsedData += ,$filteredRow  # Add the filtered row as a 2D array
        }
    }

    # Debugging: Print parsed and filtered data
    Write-Host "`n[DEBUG] Parsed Data Array (Filtered by Selected Columns):"
    Write-Host ($parsedData | Out-String)

    return $parsedData
}

function Send-KeysToQuickBooks {
    param ($dataArray)

    Add-Type -AssemblyName System.Windows.Forms

    Write-Host "`n[DEBUG] Starting data entry into QuickBooks..."

    # Check if the dataArray has data
    if (-not $dataArray -or $dataArray.Count -eq 0) {
        Write-Host "[ERROR] Data array is empty! No data to enter." -ForegroundColor Red
        return
    } else {
        Write-Host "[DEBUG] Data array contains $($dataArray.Count) rows."
    }

    # Countdown before typing starts
    Write-Host "`nPlease make sure the QuickBooks window is active and ready."
    for ($i = 5; $i -ge 1; $i--) {
        Write-Host "[DEBUG] Starting in $i seconds... (Press 'ESC' or 'Q' to cancel)"
        
        # Check for exit key before sleeping
        for ($j = 1; $j -le 10; $j++) {  # Break 1-second sleep into smaller chunks to allow fast key detection
            if ($Host.UI.RawUI.KeyAvailable) {
                $key = [System.Console]::ReadKey($true).Key
                if ($key -eq "Escape" -or $key -eq "Q") {
                    Write-Host "`n[INFO] Data entry aborted by user!" -ForegroundColor Yellow
                    return
                }
            }
            Start-Sleep -Milliseconds 100  # Sleep in small chunks to check for user input
        }
    }

    Write-Host "[DEBUG] Typing now..." -ForegroundColor Cyan

    # Iterate through each row in the 2D array
    $rowIndex = 0
    foreach ($row in $dataArray) {
        Write-Host "[DEBUG] Entering row $rowIndex..."

        $colIndex = 0
        foreach ($value in $row) {
            # Allow user to exit during typing
            if ($Host.UI.RawUI.KeyAvailable) {
                $key = [System.Console]::ReadKey($true).Key
                if ($key -eq "Escape" -or $key -eq "Q") {
                    Write-Host "`n[INFO] Data entry aborted by user!" -ForegroundColor Yellow
                    return
                }
            }


            Write-Host "[DEBUG] Typing Row ${rowIndex}, Column ${colIndex}: '${value}'"
            Start-Sleep -Milliseconds 200  # Allow time for QuickBooks to process
            
            [System.Windows.Forms.Clipboard]::Clear()  # Ensure clipboard is empty before setting new text
            Start-Sleep -Milliseconds 200
            [System.Windows.Forms.Clipboard]::SetText($value)
            Start-Sleep -Milliseconds 250
            [System.Windows.Forms.SendKeys]::SendWait("^v")  # Simulate CTRL + V (Paste)
            Start-Sleep -Milliseconds 150  # Short pause before pressing TAB
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
            $colIndex++
        }

        # Send one extra TAB to move to the next row in QuickBooks
        Start-Sleep -Milliseconds 200
        $rowIndex++
    }

    Write-Host "[DEBUG] Data entry completed successfully!" -ForegroundColor Green
}

# Step 1: Prompt user to load clipboard
$clipboardText = Get-ClipboardContent
$clipboardText = Show-TextDialog -preFill $clipboardText

# Step 2: Get clipboard contents
if ($null -eq $clipboardText) {
    Write-Host "Nothing was selected... Exiting"
    exit
}

$clipboardText = $clipboardText

# Debug: Check clipboard data
Write-Host "[DEBUG] Clipboard Content:"
Write-Host ($clipboardText | Out-String)

if (-not $clipboardText) {
    Write-Host "[ERROR] Clipboard is empty. Please copy data first and restart the script." -ForegroundColor Red
    exit
}

# Step 3: Ask user for column selection **BEFORE** parsing clipboard
Write-Host "`nPlease type the column number for each column that the following fields are in:`n"
Write-Host " Code | Description | Overdue Balance | Amount Charged `n"
Write-Host "Example: If 'Code' is in column 1, 'Description' is in column 3, etc., enter: 1,3,4,5"

$message = "`nPlease type the column number for each field:`n`n Code | Description | Overdue Balance | Amount Charged `n`n Example: If 'Code' is in column 1, 'Description' is in column 3, etc., enter: 1,3,4,5"
$userInput = Show-WideInputDialog -message $message -title "Column Selection" -defaultText "1,3,4,5"
Write-Output "You entered: $userInput"

# Capture user input safely
$columnSelection = $userInput

# Allow the user to escape
if ($columnSelection -match "^[Qq]$") {
    Write-Host "`n[INFO] Script aborted by user." -ForegroundColor Yellow
    exit
}

# Convert user input to an array of column indices (zero-based)
$selectedColumns = $columnSelection -split "," | ForEach-Object { [int]$_ - 1 }

# Debugging: Show selected columns
Write-Host "[DEBUG] User selected columns: $($selectedColumns -join '+1, ')+1"

# Step 4: Parse clipboard into a 2D array using selected columns
$dataArray = 	Convert-ClipboardData $clipboardText $selectedColumns

# Debug: Check parsed data
Write-Host "[DEBUG] Parsed Data Array:"
Write-Host ($dataArray | Out-String)

if (-not $dataArray -or $dataArray.Count -eq 0) {
    Write-Host "[ERROR] Parsed data array is empty! Check clipboard format." -ForegroundColor Red
    exit
}

# Step 5: Find available RDP windows and let the user select
$rdpWindows = Get-RDPWindows

if (0 -eq $rdpWindows.Count) {
    Write-Host "Error: No RDP windows found! Exiting..." -ForegroundColor Red
    exit
}

# Step 6: Let user select a window from the list (Option 1: Out-GridView for GUI selection)
$rdpWindows = $rdpWindows | ForEach-Object {
    [PSCustomObject]@{
        Title  = $_.Title
        Handle = $_.Handle
    }
}

# Display correctly
$rdpWindows | Format-Table -AutoSize

# Let the user select
$selectedRdpWindow = $rdpWindows | Out-GridView -Title "Select an RDP Window" -PassThru

    # If Out-GridView is unavailable, fallback to a simple console menu
    if (-not $selectedRdpWindow) {
        Write-Host "`nSelect an RDP window:"
        for ($i = 0; $i -lt $rdpWindows.Count; $i++) {
            Write-Host "$i. $($rdpWindows[$i].Title)"
        }

        $selection = Read-Host "Enter selection number"
        if ($selection -match "^\d+$" -and [int]$selection -lt $rdpWindows.Count) {
            $selectedRdpWindow = $rdpWindows[$selection]
        } else {
            Write-Host "Invalid selection. Exiting..." -ForegroundColor Red
            exit
        }
    }




# Step 7: Activate the selected RDP window
if ($selectedRdpWindow) {
    Activate-Window $selectedRdpWindow.Handle
    Write-Host "Activated: $($selectedRdpWindow.Title)" -ForegroundColor Green
} else {
    Write-Host "No window selected. Exiting..." -ForegroundColor Red
    exit
}

# Step 8: Enter Data into QuickBooks
Write-Host "Entering data into QuickBooks..." -ForegroundColor Cyan
Write-Host "[DEBUG] Checking arrayValues before sending data..."
Write-Host ($dataArray | Out-String)  # Print array content for verification
Send-KeysToQuickBooks $dataArray

Write-Host "Data entry complete!" -ForegroundColor Green
