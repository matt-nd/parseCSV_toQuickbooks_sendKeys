Add-Type -AssemblyName System.Windows.Forms

# Function to create the Kill Switch GUI
function Show-KillSwitch {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "QuickBooks Kill Switch"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true  # Keep it on top

    # Create a "Big Red Button"
    $button = New-Object System.Windows.Forms.Button
    $button.Text = "STOP QUICKBOOKS ENTRY"
    $button.BackColor = [System.Drawing.Color]::Red
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $button.Size = New-Object System.Drawing.Size(250,50)
    $button.Location = New-Object System.Drawing.Point(20,20)

    # Button Click Event - Kill QuickBooks Automation Process
    $button.Add_Click({
        $processName = "powershell"
        
        # Confirm before killing process
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to stop QuickBooks automation?", 
            "Confirm Kill", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Find all PowerShell processes
            $powershellProcesses = Get-Process | Where-Object { $_.ProcessName -eq $processName }

            if ($powershellProcesses) {
                Write-Host "[INFO] Killing QuickBooks automation process..."
                $powershellProcesses | Stop-Process -Force
            } else {
                [System.Windows.Forms.MessageBox]::Show("No running QuickBooks automation found!", 
                    "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    })

    # Add button to form
    $form.Controls.Add($button)

    # Show the form (Blocking UI Thread)
    $form.ShowDialog()
}

# Run the Kill Switch GUI in a New PowerShell Process (Non-Blocking)
Show-KillSwitch