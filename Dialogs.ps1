Add-Type -AssemblyName PresentationFramework

# Define XAML as a string
$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Large Input Box" Height="400" Width="500" WindowStartupLocation="CenterScreen">
    <Grid>
        <RichTextBox Name="InputBox" VerticalScrollBarVisibility="Auto"
                     AcceptsReturn="True" Margin="10,10,10,50"/>
        <Button Content="OK" Name="OKButton" Width="80" Height="30"
                HorizontalAlignment="Center" VerticalAlignment="Bottom" Margin="10"/>
    </Grid>
</Window>
"@

# Function to show dialog with a prefilled text
function Show-TextDialog {
    param([string]$preFill)

    # Load XAML
    $reader = (New-Object System.Xml.XmlNodeReader ([xml]$XAML))
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Get references to UI elements
    $InputBox = $window.FindName("InputBox")
    $OKButton = $window.FindName("OKButton")

    # Set initial text
    $flowDoc = New-Object Windows.Documents.FlowDocument
    $paragraph = New-Object Windows.Documents.Paragraph
    $paragraph.Inlines.Add($preFill)
    $flowDoc.Blocks.Add($paragraph)
    $InputBox.Document = $flowDoc

    # Default return value (handles "X" close scenario)
    $script:finalText = $null 

    # Handle OK button click
    $OKButton.Add_Click({
        $textRange = New-Object Windows.Documents.TextRange ($InputBox.Document.ContentStart, $InputBox.Document.ContentEnd)
        $script:finalText = $textRange.Text.Trim()
        $window.DialogResult = $true  # Closes the window with success
    })

    # Show Window
    $result = $window.ShowDialog() 

    # Return result based on how the window was closed
    if ($result -eq $true) {
        return $script:finalText  # OK was clicked
    } else {
        return $null  # Window was closed with "X"
    }
}

## Example usage
# $preFilledText = "This is a prefilled large text input box.`r`nYou can edit this text as needed."
# $userInput = Show-TextDialog -preFill $preFilledText

## Output the user-entered text
#if ($null -eq $variable) {
#    Write-Output "User entered:`n$userInput"
#} else {
#    Write-Output "Input was canceled (user clicked 'X')."
#}

Add-Type -AssemblyName System.Windows.Forms

function Show-WideInputDialog {
    param([string]$message, [string]$title, [string]$defaultText)

    # Create a form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size(500, 200)  # WIDER WIDTH
    $form.StartPosition = "CenterScreen"

    # Create label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(20, 20)

    # Create input box
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Size = New-Object System.Drawing.Size(400, 30)  # Wider TextBox
    $textbox.Location = New-Object System.Drawing.Point(25, 105)
    $textbox.Text = $defaultText  # Set default text

    # Create OK button
    $buttonOK = New-Object System.Windows.Forms.Button
    $buttonOK.Text = "OK"
    $buttonOK.Location = New-Object System.Drawing.Point(25, 135)

    # Add event to close form and return input
    $buttonOK.Add_Click({
        $script:UserInput = $textbox.Text
        $form.Close()
    })

    # Add controls to form
    $form.Controls.Add($label)
    $form.Controls.Add($textbox)
    $form.Controls.Add($buttonOK)

    # Show form
    $form.ShowDialog() | Out-Null
    return $script:UserInput
}

<# ## Example usage
$message = "`nPlease type the column number for each field:`n`n Code | Description | Overdue Balance | Amount Charged `n`n Example: If 'Code' is in column 1, 'Description' is in column 3, etc., enter: 1,3,4,5"
$userInput = Show-WideInputDialog -message $message -title "Column Selection"
#Write-Output "You entered: $userInput"
 #>