# Start the Kill Switch GUI
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"C:\Users\matt\Desktop\InvoiceInput\KillSwitch.ps1`"" 

# Start QuickBooks automation with proper logging
Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -NoExit -File `"C:\Users\matt\Desktop\InvoiceInput\PasteInvoices.ps1`"" -WindowStyle Normal