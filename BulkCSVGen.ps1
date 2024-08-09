Add-Type -AssemblyName PresentationFramework

# Create the main window
$mainWindow = New-Object system.windows.window
$mainWindow.Title = "CSV Gen"
$mainWindow.Width = 600
$mainWindow.Height = 300

# Create a Grid layout
$grid = New-Object system.windows.controls.grid

# Define rows and columns
for ($i = 0; $i -lt 6; $i++) {
    $rowDef = New-Object system.windows.controls.rowdefinition
    $grid.RowDefinitions.Add($rowDef)
}

for ($i = 0; $i -lt 2; $i++) {
    $colDef = New-Object system.windows.controls.columndefinition
    $grid.ColumnDefinitions.Add($colDef)
}

# Create labels and textboxes for first name, last name, and email
$firstNameLabel = New-Object system.windows.controls.label
$firstNameLabel.Content = "First Name:"
$firstNameLabel.Margin = "10"
$grid.Children.Add($firstNameLabel)
[system.windows.controls.grid]::SetRow($firstNameLabel, 0)
[system.windows.controls.grid]::SetColumn($firstNameLabel, 0)

$firstNameTextBox = New-Object system.windows.controls.textbox
$firstNameTextBox.Margin = "10"
$firstNameTextBox.Width = 250
$grid.Children.Add($firstNameTextBox)
[system.windows.controls.grid]::SetRow($firstNameTextBox, 0)
[system.windows.controls.grid]::SetColumn($firstNameTextBox, 1)

$lastNameLabel = New-Object system.windows.controls.label
$lastNameLabel.Content = "Last Name:"
$lastNameLabel.Margin = "10"
$grid.Children.Add($lastNameLabel)
[system.windows.controls.grid]::SetRow($lastNameLabel, 1)
[system.windows.controls.grid]::SetColumn($lastNameLabel, 0)

$lastNameTextBox = New-Object system.windows.controls.textbox
$lastNameTextBox.Margin = "10"
$lastNameTextBox.Width = 250
$grid.Children.Add($lastNameTextBox)
[system.windows.controls.grid]::SetRow($lastNameTextBox, 1)
[system.windows.controls.grid]::SetColumn($lastNameTextBox, 1)

$emailLabel = New-Object system.windows.controls.label
$emailLabel.Content = "Email:"
$emailLabel.Margin = "10"
$grid.Children.Add($emailLabel)
[system.windows.controls.grid]::SetRow($emailLabel, 2)
[system.windows.controls.grid]::SetColumn($emailLabel, 0)

$emailTextBox = New-Object system.windows.controls.textbox
$emailTextBox.Margin = "10"
$emailTextBox.Width = 250
$grid.Children.Add($emailTextBox)
[system.windows.controls.grid]::SetRow($emailTextBox, 2)
[system.windows.controls.grid]::SetColumn($emailTextBox, 1)

# Create a button to add the user data
$addButton = New-Object system.windows.controls.button
$addButton.Content = "Add Person"
$addButton.Margin = "10"
$addButton.Width = 100
$grid.Children.Add($addButton)
[system.windows.controls.grid]::SetRow($addButton, 3)
[system.windows.controls.grid]::SetColumn($addButton, 0)

# Create a button to submit the data and generate the CSV
$submitButton = New-Object system.windows.controls.button
$submitButton.Content = "Submit"
$submitButton.Margin = "10"
$submitButton.Width = 100
$grid.Children.Add($submitButton)
[system.windows.controls.grid]::SetRow($submitButton, 3)
[system.windows.controls.grid]::SetColumn($submitButton, 1)

# Create a list to store the user data
$userData = [System.Collections.ArrayList]@()

# Add event handler for the Add button
$addButton.Add_Click({
    $firstName = $firstNameTextBox.Text
    $lastName = $lastNameTextBox.Text
    $email = $emailTextBox.Text

    # Create a custom object to store the data
    $user = [PSCustomObject]@{
        Employee  = ""
        FIRST_NAME = $firstName
        LAST_NAME  = $lastName
        EMAIL     = $email
    }

    # Add the user to the list
    [void]$userData.Add($user)

    # Clear the textboxes for new entry
    $firstNameTextBox.Clear()
    $lastNameTextBox.Clear()
    $emailTextBox.Clear()

    # Debugging output to confirm data is added
    Write-Host "User added: $($user | Out-String)"
})

# Add event handler for the Submit button
$submitButton.Add_Click({
    # Check if there is any data to export
    if ($userData.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No user data to export.")
        return
    }

    $filePath = [System.IO.Path]::Combine([System.Environment]::GetFolderPath('Desktop'), 'batch.csv')

    try {
        # Export the data to CSV
        $userData | Select-Object Employee, FIRST_NAME, LAST_NAME, EMAIL | Export-Csv -Path $filePath -NoTypeInformation
        [System.Windows.MessageBox]::Show("CSV file created at $filePath")
        # Debugging output to confirm data export
        Write-Host "CSV file created at $filePath with the following data:"
        $userData | Format-Table | Out-String | Write-Host
    } catch {
        [System.Windows.MessageBox]::Show("Failed to create CSV file. Error: $($_.Exception.Message)")
        Write-Host "Failed to create CSV file. Error: $($_.Exception.Message)"
    }
})

# Set the content of the main window and show it
$mainWindow.Content = $grid
$mainWindow.ShowDialog()
