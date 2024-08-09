# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define API key and base URL
$apiKey = "your_api_key"
$apiUrl = "https://apps.esignlive.eu/api"

# Define headers for authentication
$headers = @{
    "Authorization" = "Basic $apiKey"
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "OneSpan Sign Bulk Send"
$form.Size = New-Object System.Drawing.Size(500, 200)
$form.StartPosition = "CenterScreen"

# Create label and textbox for template ID
$labelTemplateId = New-Object System.Windows.Forms.Label
$labelTemplateId.Text = "Template ID:"
$labelTemplateId.Location = New-Object System.Drawing.Point(10, 20)
$labelTemplateId.Size = New-Object System.Drawing.Size(130, 20)
$form.Controls.Add($labelTemplateId)

$textBoxTemplateId = New-Object System.Windows.Forms.TextBox
$textBoxTemplateId.Location = New-Object System.Drawing.Point(150, 20)
$textBoxTemplateId.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxTemplateId)

# Create label and textbox for CSV file path
$labelCsvPath = New-Object System.Windows.Forms.Label
$labelCsvPath.Text = "CSV File Path:"
$labelCsvPath.Location = New-Object System.Drawing.Point(10, 60)
$labelCsvPath.Size = New-Object System.Drawing.Size(130, 20)
$form.Controls.Add($labelCsvPath)

$textBoxCsvPath = New-Object System.Windows.Forms.TextBox
$textBoxCsvPath.Location = New-Object System.Drawing.Point(150, 60)
$textBoxCsvPath.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxCsvPath)

# Create a button to browse CSV file
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Browse..."
$buttonBrowse.Location = New-Object System.Drawing.Point(360, 60)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 30)
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxCsvPath.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowse)

# Create a button to submit the form
$buttonSubmit = New-Object System.Windows.Forms.Button
$buttonSubmit.Text = "Submit"
$buttonSubmit.Location = New-Object System.Drawing.Point(150, 100)
$buttonSubmit.Size = New-Object System.Drawing.Size(80, 30)
$buttonSubmit.Add_Click({
    $templateId = $textBoxTemplateId.Text
    $csvFilePath = $textBoxCsvPath.Text

    if (-not (Test-Path -Path $csvFilePath)) {
        [System.Windows.Forms.MessageBox]::Show("CSV file not found at path: $csvFilePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Function to test the connection
    try {
        $testUrl = "$apiUrl/packages"
        $response = Invoke-RestMethod -Uri $testUrl -Headers $headers -Method Get
        [System.Windows.Forms.MessageBox]::Show("Connection Successful.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Connection Failed. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Function to upload CSV file and bulk send transactions
    try {
        $bulkSendUrl = "$apiUrl/packages/$templateId/bulk_send"
        
        # Prepare the multipart/form-data body
        $boundary = [System.Guid]::NewGuid().ToString()
        $headers["Content-Type"] = "multipart/form-data; boundary=$boundary"
        $fileContent = [System.IO.File]::ReadAllBytes($csvFilePath)
        $fileName = [System.IO.Path]::GetFileName($csvFilePath)
        $fileData = [System.Text.Encoding]::UTF8.GetString($fileContent)
        $bodyLines = @(
            "--$boundary",
            "Content-Disposition: form-data; name=`"file`"; filename=`"$fileName`"",
            "Content-Type: text/csv",
            "",
            $fileData,
            "--$boundary--"
        )
        $body = [System.Text.Encoding]::UTF8.GetBytes($bodyLines -join "`r`n")

        # Invoke the API call to bulk send transactions
        $response = Invoke-RestMethod -Uri $bulkSendUrl -Headers $headers -Method Post -Body $body
        [System.Windows.Forms.MessageBox]::Show("Bulk Send Successful.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Bulk Send Failed. Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$form.Controls.Add($buttonSubmit)

# Show the form
[void]$form.ShowDialog()
