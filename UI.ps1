###########################################################################################################
#     Initialize Form
###########################################################################################################

function Show-UserForm {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Add Directory User"
    $form.Size = New-Object System.Drawing.Size(420, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle
    $form.MaximizeBox = $false
    $form.MinimizeBox = $true

    $font = New-Object System.Drawing.Font("Arial", 10)

###########################################################################################################
#     OU Dropdown Setup
###########################################################################################################

    # Panel for inputs
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(400, 450)
    $panel.Location = New-Object System.Drawing.Point(10, 10)

    $labels = @("First Name:", "Initial:", "Last Name:", "Username:", "Email:", "Password:", "User ID:", "Department:", "Job Title:", "Company:","Class:")
    $textBoxes = @()

    # OU label and dropdown (placed above other fields)
    $ouLabel = New-Object System.Windows.Forms.Label
    $ouLabel.Text = "OU:"
    $ouLabel.Location = New-Object System.Drawing.Point(10, 0)
    $ouLabel.Size = New-Object System.Drawing.Size(120, 25)
    $ouLabel.Font = $font
    $panel.Controls.Add($ouLabel)

    $ouDropdown = New-Object System.Windows.Forms.ComboBox
    $ouDropdown.Location = New-Object System.Drawing.Point(140, 0)
    $ouDropdown.Size = New-Object System.Drawing.Size(220, 25)
    $ouDropdown.Font = $font
    $ouDropdown.DropDownStyle = "DropDownList"
    $ouDropdown.Items.AddRange(@("Staff", "Students", "Faculty", "EmailOnly", "BulkStudents"))
    $ouDropdown.SelectedIndex = 0
    $panel.Controls.Add($ouDropdown)

###########################################################################################################
#     Input Fields Setup
###########################################################################################################

    for ($i = 0; $i -lt $labels.Count; $i++) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $labels[$i]
        $label.Location = New-Object System.Drawing.Point(10, (35 * ($i + 1)))
        $label.Size = New-Object System.Drawing.Size(120, 25)
        $label.Font = $font
        $panel.Controls.Add($label)

        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(140, (35 * ($i + 1)))
        $textBox.Size = New-Object System.Drawing.Size(220, 25)
        $textBox.Font = $font
        $textBox.ForeColor = "Black"
        $textBox.BackColor = "White"

        if ($labels[$i] -eq "Password:") {
            $textBox.Text = "ChangeMe123!"  # default placeholder password
        }

        $panel.Controls.Add($textBox)
        $textBoxes += $textBox
    }

    # Default selected Company to match default selected OU on form load
    $textBoxes[9].Text = "Staff with Office PC"

    # Declare class label / textbox to hide initially
    $classLabel = $panel.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Text -eq "Class:" }
    $classTextBox = $textBoxes[10]

    $classLabel.Visible = $false
    $classTextBox.Visible = $false

###########################################################################################################
#     Generate Username and Email
###########################################################################################################

   $generateUsername = {
        if ($textBoxes[0].Text -ne "" -and $textBoxes[2].Text -ne "") {
            $first = $textBoxes[0].Text.Trim().ToLower()
            $last = $textBoxes[2].Text.Trim().ToLower()
            $ou = $ouDropdown.SelectedItem

            if ($ou -eq "Students") {
                $username = "$first.$last"
                $email = "$username@student.example.edu"
            } else {
                $username = "$($first.Substring(0,1))$last"
                $email = "$username@example.edu"
            }

            $textBoxes[3].Text = $username
            $textBoxes[4].Text = $email
        }
    }

    $textBoxes[0].Add_TextChanged($generateUsername)
    $textBoxes[2].Add_TextChanged($generateUsername)
    $ouDropdown.Add_SelectedIndexChanged($generateUsername)

###########################################################################################################
#     Image & Label for Bulk Import Example (optional)
###########################################################################################################

    $imagePath = Join-Path -Path $PSScriptRoot -ChildPath "BulkImportExample.png"
    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(380, 150)
    $pictureBox.Location = New-Object System.Drawing.Point(20, 420)
    $pictureBox.SizeMode = "Zoom"
    $pictureBox.BorderStyle = "FixedSingle"
    $pictureBox.Visible = $false

    if (Test-Path $imagePath) {
        $pictureBox.Image = [System.Drawing.Image]::FromFile($imagePath)
    } else {
        $pictureBox.BackColor = "LightGray"
    }

    $exampleLabel = New-Object System.Windows.Forms.Label
    $exampleLabel.Text = "Example Import Format Image"
    $exampleLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Underline)
    $exampleLabel.ForeColor = "Blue"
    $exampleLabel.Location = New-Object System.Drawing.Point(20, 580)
    $exampleLabel.AutoSize = $true
    $exampleLabel.Cursor = [System.Windows.Forms.Cursors]::Hand
    $exampleLabel.Visible = $false

    $exampleLabel.Add_Click({
        $imageForm = New-Object System.Windows.Forms.Form
        $imageForm.Text = "Example Screenshot"
        $imageForm.Size = New-Object System.Drawing.Size(800, 600)
        $imageForm.StartPosition = "CenterScreen"

        $imgBox = New-Object System.Windows.Forms.PictureBox
        $imgBox.Dock = "Fill"
        $imgBox.SizeMode = "Zoom"
        $imgBox.Image = [System.Drawing.Image]::FromFile($imagePath)

        $imageForm.Controls.Add($imgBox)
        $imageForm.ShowDialog()
    })

###########################################################################################################
#     OU Dropdown Behavior
###########################################################################################################

    $ouDropdown.Add_SelectedIndexChanged({
        $selected = $ouDropdown.SelectedItem
        switch ($selected) {
            "EmailOnly"     { $textBoxes[9].Text = "Email Only User" }
            "Staff"         { $textBoxes[9].Text = "Staff with Office PC" }
            "Students"      { $textBoxes[9].Text = "Student" }
            "Faculty"       { $textBoxes[9].Text = "Faculty" }
            "BulkStudents"  { $textBoxes[9].Text = "Student Bulk Add" }
            default        { $textBoxes[9].Text = "" }
        }

        $isBulk = ($selected -eq "BulkStudents")
        $isStudent = ($selected -eq "Students")

        foreach ($control in $panel.Controls) {
            if ($control -is [System.Windows.Forms.Label] -or $control -is [System.Windows.Forms.TextBox]) {
                if ($control -ne $ouLabel -and $control -ne $ouDropdown) {
                    $control.Visible = -not $isBulk
                }
            }
        }

        $pictureBox.Visible = $isBulk
        $exampleLabel.Visible = $isBulk

        $classLabel.Visible = $isStudent
        $classTextBox.Visible = $isStudent

        if ($isBulk) {
            $button.Text = "Next"
        } else {
            $button.Text = "Add User"
        }
    })

###########################################################################################################
#     Button Setup and Click Event
###########################################################################################################

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Add User"
    $button.Font = $font
    $button.Size = New-Object System.Drawing.Size(100, 30)
    $button.Location = New-Object System.Drawing.Point(150, 620)
    $button.FlatStyle = "Flat"
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $button.ForeColor = [System.Drawing.Color]::White
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand

    $button.Add_Click({
        if ($ouDropdown.SelectedItem -eq "BulkStudents") {
            . "$PSScriptRoot\BulkStudentImportUI.ps1"
            $form.Hide()
            $importResult = Show-MassStudentImportForm

            if ($importResult -eq [System.Windows.Forms.DialogResult]::OK) {
                $form.Show()
            } elseif ($importResult -eq [System.Windows.Forms.DialogResult]::Retry) {
                $form.Show()
            } else {
                $form.Close()
            }

            return
        }

        $userData = @(
            $textBoxes[0].Text
            $textBoxes[1].Text
            $textBoxes[2].Text
            $textBoxes[3].Text
            $textBoxes[4].Text
            $textBoxes[5].Text
            $textBoxes[6].Text
            $textBoxes[7].Text
            $textBoxes[8].Text
            $ouDropdown.SelectedItem
            $textBoxes[9].Text
        )

        if ($userData[0] -eq "" -or $userData[2] -eq "" -or $userData[4] -eq "" -or $userData[5] -eq "") {
            [System.Windows.Forms.MessageBox]::Show("All fields except Initial are required!", "Error", "OK", "Error")
            return
        }
        if ($classTextBox.Visible -and $classTextBox.Text -ne "") {
            $userData += $classTextBox.Text
        }

        $result = Add-ADUserSafely -UserData $userData
        [System.Windows.Forms.MessageBox]::Show($form, $result, "Result", "OK", "Information")
    })

###########################################################################################################

    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "Clear All"
    $clearButton.Font = $font
    $clearButton.Size = New-Object System.Drawing.Size(100, 30)
    $clearButton.Location = New-Object System.Drawing.Point(150, 485)
    $clearButton.FlatStyle = "Flat"
    $clearButton.BackColor = [System.Drawing.Color]::FromArgb(225, 225, 225)
    $clearButton.ForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $clearButton.Cursor = [System.Windows.Forms.Cursors]::Hand

    $clearButton.Add_Click({
        foreach ($tb in $textBoxes) {
            $tb.Text = ""
        }
        $ouDropdown.SelectedIndex = 0
        $textBoxes[9].Text = "Staff with Office PC"
        $textBoxes[5].Text = "ChangeMe123!"
        $classLabel.Visible = $false
        $classTextBox.Visible = $false
        $classTextBox.Text = ""
    })

###########################################################################################################

    $toggleConsoleButton = New-Object System.Windows.Forms.Button
    $toggleConsoleButton.Size = New-Object System.Drawing.Size(30, 30)
    $toggleConsoleButton.Location = New-Object System.Drawing.Point(0, 660)
    $toggleConsoleButton.FlatStyle = "Flat"
    $toggleConsoleButton.BackColor = [System.Drawing.Color]::Transparent
    $toggleConsoleButton.FlatAppearance.BorderSize = 0
    $toggleConsoleButton.Cursor = [System.Windows.Forms.Cursors]::Hand
    $toggleConsoleButton.Text = [char]0x2328   # Unicode keyboard icon
    $toggleConsoleButton.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 16)
    $toggleConsoleButton.TabStop = $false

    $toolTip = New-Object System.Windows.Forms.ToolTip
    $toolTip.SetToolTip($toggleConsoleButton, "Toggle Console Window")

    $toggleConsoleButton.Add_Click({
        $isVisible = [Win32]::IsWindowVisible($global:consoleHandle)
        if ($isVisible) {
            [Win32]::ShowWindow($global:consoleHandle, [Win32]::SW_HIDE) | Out-Null
        } else {
            [Win32]::ShowWindow($global:consoleHandle, [Win32]::SW_SHOW) | Out-Null
        }
    })

###########################################################################################################
#     Final Layout and Form Logic
###########################################################################################################

    $form.TopMost = $true
    $form.Controls.Add($panel)
    $form.Controls.Add($pictureBox)
    $form.Controls.Add($exampleLabel)
    $form.Controls.Add($clearButton)
    $form.Controls.Add($button)
    $form.Controls.Add($toggleConsoleButton)

    $form.Add_FormClosing({
        Write-Host "User closed the form. Exiting..."
        return "Closed"
    })

    $form.ShowDialog()
}
