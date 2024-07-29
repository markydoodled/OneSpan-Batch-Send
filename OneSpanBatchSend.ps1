Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create a new transaction based on a template and send it to multiple recipients
function CreateAndSendBatchTransaction {
    param (
        [string]$templateID,
        [string[]]$recipientEmails
    )

    # Define the OneSpan API endpoint and the API key (replace with your actual API key)
    $apiUrl = "https://api.onespan.com/api/packages"
    $apiKey = "YOUR_API_KEY"

    # Create the recipients array
    $recipients = @()
    foreach ($email in $recipientEmails) {
        $recipients += @{
            "email" = $email
            "role" = "SIGNER"
        }
    }

    # Define the request body for creating a new transaction based on a template
    $requestBody = @{
        "templateId" = $templateID
        "recipients" = $recipients
    } | ConvertTo-Json -Depth 3

    # Make the API call to create a new transaction
    try {
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers @{ "Authorization" = "Bearer $apiKey"; "Content-Type" = "application/json" } -Body $requestBody

        # Display the transaction ID
        [System.Windows.Forms.MessageBox]::Show("Transaction Created Successfully. Transaction ID: $($response.id)")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_")
    }
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create And Send Batch OneSpan Transaction"
$form.Size = New-Object System.Drawing.Size(500, 300)

# Create the template ID label and textbox
$templateIDLabel = New-Object System.Windows.Forms.Label
$templateIDLabel.Text = "Template ID:"
$templateIDLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($templateIDLabel)

$templateIDTextbox = New-Object System.Windows.Forms.TextBox
$templateIDTextbox.Location = New-Object System.Drawing.Point(100, 20)
$templateIDTextbox.Size = New-Object System.Drawing.Size(350, 20)
$form.Controls.Add($templateIDTextbox)

# Create the recipient emails label and textbox
$recipientEmailsLabel = New-Object System.Windows.Forms.Label
$recipientEmailsLabel.Text = "Recipient Emails (Comma-Separated):"
$recipientEmailsLabel.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($recipientEmailsLabel)

$recipientEmailsTextbox = New-Object System.Windows.Forms.TextBox
$recipientEmailsTextbox.Location = New-Object System.Drawing.Point(10, 90)
$recipientEmailsTextbox.Size = New-Object System.Drawing.Size(450, 100)
$recipientEmailsTextbox.Multiline = $true
$form.Controls.Add($recipientEmailsTextbox)

# Create the submit button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Create And Send"
$submitButton.Location = New-Object System.Drawing.Point(10, 200)
$submitButton.Size = New-Object System.Drawing.Size(150, 30)
$submitButton.Add_Click({
    $templateID = $templateIDTextbox.Text
    $recipientEmails = $recipientEmailsTextbox.Text -split "\s*,\s*"
    CreateAndSendBatchTransaction -templateID $templateID -recipientEmails $recipientEmails
})
$form.Controls.Add($submitButton)

# Show the form
[void]$form.ShowDialog()
