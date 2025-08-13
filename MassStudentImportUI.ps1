###########################################################################################################
#     Initialize Function & Form & Font Setup
###########################################################################################################

function Show-MassStudentImportForm {
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Mass Student Import"
$form.Size = New-Object System.Drawing.Size(460, 300)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$font = New-Object System.Drawing.Font("Arial", 10)

###########################################################################################################
#     Labels & Buttons
###########################################################################################################
# Top-left back button
$backButton = New-Object System.Windows.Forms.Button
$backButton.Text = "<= Back"
$backButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$backButton.Size = New-Object System.Drawing.Size(80, 30)
$backButton.Location = New-Object System.Drawing.Point(10, 10)
$backButton.FlatStyle = "Flat"
$backButton.BackColor = "WhiteSmoke"
$backButton.ForeColor = "Black"
$backButton.Cursor = [System.Windows.Forms.Cursors]::Hand

# File path label
$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Text = "No file selected"
$fileLabel.AutoSize = $false
$fileLabel.Size = New-Object System.Drawing.Size(380, 45)
$fileLabel.Location = New-Object System.Drawing.Point(30, 50)
$fileLabel.Font = $font

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = ""
$statusLabel.Size = New-Object System.Drawing.Size(380, 25)
$statusLabel.Location = New-Object System.Drawing.Point(30, 180)
$statusLabel.Font = $font
$statusLabel.ForeColor = "Blue"

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Size = New-Object System.Drawing.Size(380, 20)
$progressBar.Location = New-Object System.Drawing.Point(30, 210)
$progressBar.Visible = $false

# Select File Button
$selectFileButton = New-Object System.Windows.Forms.Button
$selectFileButton.Text = "Select CSV File"
$selectFileButton.Size = New-Object System.Drawing.Size(150, 30)
$selectFileButton.Location = New-Object System.Drawing.Point(30, 100)
$selectFileButton.Font = $font
$selectFileButton.FlatStyle = "Flat"
$selectFileButton.BackColor = "WhiteSmoke"
$selectFileButton.ForeColor = "Black"
$selectFileButton.Cursor = [System.Windows.Forms.Cursors]::Hand

# Submit Button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Size = New-Object System.Drawing.Size(150, 30)
$submitButton.Location = New-Object System.Drawing.Point(30, 140)
$submitButton.Font = $font
$submitButton.FlatStyle = "Flat"
$submitButton.ForeColor = "DarkGreen"
$submitButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$submitButton.BackColor = "LightGreen"
$submitButton.Visible = $false

# Store selected file path
$script:selectedFilePath = $null

###########################################################################################################
#     File Select Logic
###########################################################################################################

# File select logic
$selectFileButton.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $dialog.Title = "Select Student CSV File"

    if ($dialog.ShowDialog() -eq "OK") {
        $script:selectedFilePath = $dialog.FileName
        $fileLabel.Text = "Selected: $selectedFilePath"
        $submitButton.Visible = $true
        $statusLabel.Text = ""
        $progressBar.Visible = $false
    }
})

###########################################################################################################
#     Button Logic
###########################################################################################################

# Submit logic: Run external logic script
$submitButton.Add_Click({
    if (-not $selectedFilePath) {
        [System.Windows.Forms.MessageBox]::Show("No file selected.", "Error", "OK", "Error")
        return
    }

    try {
        $progressBar.Visible = $true
        $progressBar.Value = 5
        $statusLabel.Text = "Starting import..."

        Start-Sleep -Milliseconds 300

        $progressBar.Value = 20
        $statusLabel.Text = "Validating file..."
        Start-Sleep -Milliseconds 200

        if (-not (Test-Path $selectedFilePath)) {
            throw "File not found: $selectedFilePath"
        }

        if ([System.IO.Path]::GetExtension($selectedFilePath) -ne ".csv") {
            throw "Selected file must be a CSV."
        }

        $progressBar.Value = 40
        $statusLabel.Text = "Running import script..."

        # Run the logic script and capture result
        $result = & "$PSScriptRoot\MassStudentImportLogic.ps1" -CsvPath $selectedFilePath

        $progressBar.Value = 80

        if ($result -eq "SUCCESS") {
            $progressBar.Value = 100
            $statusLabel.Text = "Import complete!"
            [System.Windows.Forms.MessageBox]::Show("Student import completed successfully!", "Success", "OK", "Information")
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
	    $form.Close()
        }
        else {
            throw "The logic script did not report success. Message: $result"
        }
    }
    catch {
        $statusLabel.Text = "Failed."
        
	$form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    }
})

#back button logic: go back to UI.ps1 form
$backButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.Close()
})


###########################################################################################################
#     Form Elements (controls) & Layout
###########################################################################################################

# Add controls
$form.Controls.Add($backButton)
$form.Controls.Add($fileLabel)
$form.Controls.Add($statusLabel)
$form.Controls.Add($progressBar)
$form.Controls.Add($selectFileButton)
$form.Controls.Add($submitButton)

# Show the form
$result = $form.ShowDialog()
return $result
}
