function Add-ADUserSafely {
    param (
        [array]$UserData
    )

    # Extract values
    $FirstName      = $UserData[0]
    $Initial        = $UserData[1]
    $LastName       = $UserData[2]
    $SamAccountName = $UserData[3]  # Username
    $Password       = $UserData[5]
    $EmployeeID     = $UserData[6]
    $Department     = $UserData[7]
    $JobTitle       = $UserData[8]
    $SelectedOU     = $UserData[9]
    $Company        = $UserData[10]
    $Class          = $UserData[11]

    # Set default properties
    $Email = "$SamAccountName@example.edu"
    $HomeDirectory  = "\\server\users\$SamAccountName"
    $ScriptPath = "default_login_script.bat"

    # OU-specific overrides
    switch ($SelectedOU) {
        "EMAIL_ONLY" { $ouPath = "OU=EmailOnly,OU=Staff,DC=example,DC=edu" }
        "STAFF"      { $ouPath = "OU=Users,OU=Staff,DC=example,DC=edu" }
        "STUDENTS"   {
            $ouPath = "OU=Users,OU=Students,DC=example,DC=edu"
            $Email = "$($FirstName.ToLower()).$($LastName.ToLower())@students.example.edu"
            $HomeDirectory = "\\studentserver\$Class\$SamAccountName"
            $ScriptPath = "default_student_login_script.bat"
        }
        "FACULTY"    { $ouPath = "OU=Users,OU=Faculty,DC=example,DC=edu" }
        default     { $ouPath = "OU=Users,OU=Staff,DC=example,DC=edu" }
    }

    # Convert password securely
    $securePassword = ConvertTo-SecureString -AsPlainText $Password -Force

    # User properties hash
    $userParams = @{
        Name              = $SamAccountName
        GivenName         = $FirstName
        Initials          = $Initial
        Surname           = $LastName
        DisplayName       = "$LastName, $FirstName"
        SamAccountName    = $SamAccountName
        UserPrincipalName = $Email
        EmailAddress      = $Email
        Department        = $Department
        Title             = $JobTitle
        Description       = $JobTitle
        Company           = $Company
        EmployeeID        = $EmployeeID
        ScriptPath        = $ScriptPath
        HomeDirectory     = $HomeDirectory
        HomeDrive         = "X:"
        Path              = $ouPath
        AccountPassword   = $securePassword
        Enabled           = $true
    }

    try {
        # Create user account
        New-ADUser @userParams -Server $env:USERDNSDOMAIN

        # Update additional attributes
        Set-ADUser -Identity $SamAccountName -Replace @{
            Pager        = $EmployeeID
            MailNickname = $SamAccountName
        }

        # Set 'Password Never Expires'
        $user = Get-ADUser -Identity $SamAccountName -Properties UserAccountControl
        $userAccountControl = $user.UserAccountControl -bor 0x10000
        Set-ADUser -Identity $SamAccountName -Replace @{userAccountControl = $userAccountControl}

        # Add user to groups (generic placeholders)
        $groups = @("All_Staff", "All_Admins", "Office365_Faculty", "STAFF_NETWORK", "STAFF_WIFI", "VPN_Access")
        foreach ($group in $groups) {
            Add-ADGroupMember -Identity $group -Members $SamAccountName
        }

        Write-Log "User $SamAccountName added successfully with email $Email"

        # Send welcome email for certain OUs
        $allowedOUs = @("STAFF", "FACULTY", "EMAIL_ONLY")

        if ($allowedOUs -contains $SelectedOU) {
            $htmlBody = @"
<html>
<body style='font-family: Arial, sans-serif; font-size: 14px; color: #333;'>
<p>Hello $FirstName,</p>
<p>Your account has been created. Details below:</p>
<table cellpadding='6' cellspacing='0' border='0' style='border-collapse: collapse;'>
<tr><td><b>Username:</b></td><td>$SamAccountName</td></tr>
<tr><td><b>Password:</b></td><td>Welcome123!</td></tr>
<tr><td><b>Email Address:</b></td><td>$Email</td></tr>
</table>
<p><b>Important:</b> Email sync may take up to an hour.</p>
<h3>First Login Instructions:</h3>
<ol>
<li>Sign in on a campus computer.</li>
<li>Change your password via <b>Ctrl + Alt + Delete</b>.</li>
</ol>
<h3>Portal Access:</h3>
<ul>
<li>Visit <a href='https://portal.example.edu'>portal.example.edu</a> and log in with your credentials.</li>
</ul>
<h3>Email & MFA Setup:</h3>
<ul>
<li>Refer to the IT Helpdesk section for setup instructions.</li>
</ul>
<h3>Wi-Fi Setup:</h3>
<ul>
<li>Connect to campus Wi-Fi using your username and password.</li>
</ul>
<p>Contact IT Support if you have questions.</p>
<p>Thanks,<br>IT Support Team</p>
</body>
</html>
"@

            $outlook = New-Object -ComObject Outlook.Application
            $mail = $outlook.CreateItem(0)
            $mail.To = $Email
            $mail.Subject = "$FirstName $LastName - Account Information"
            $mail.HTMLBody = $htmlBody
            $mail.Display()
            Write-Log "Welcome email draft created for $SamAccountName"
        }
        else {
            Write-Log "No email sent for OU: $SelectedOU"
        }

        return "User $SamAccountName added successfully!"
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log "Error adding user $SamAccountName: $errorMessage" -Level Error
        return "Error: Failed to add user $SamAccountName."
    }
}
