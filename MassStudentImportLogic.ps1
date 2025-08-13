param (
    [string]$CsvPath
)

function Add-Mass-ADStudentSafely {
    param (
        [array]$UserData
    )

    $Password   = $UserData[0]
    $FirstName  = $UserData[1]
    $LastName   = $UserData[2]
    $StudentID  = $UserData[3]
    $Initial    = $UserData[4]
    $Class      = $UserData[5]

    $SamAccountName = "${FirstName}.${LastName}"
    $Email          = "${SamAccountName}@students.example.edu"
    $DisplayName    = "$LastName, $FirstName"
    $ouPath         = "OU=Users,OU=Students,DC=example,DC=edu"
    $HomeDirectory  = "\\server_placeholder\${Class}\${SamAccountName}"

    $userParams = @{
        Name              = $SamAccountName
        GivenName         = $FirstName
        Initials          = $Initial
        Surname           = $LastName
        DisplayName       = $DisplayName
        SamAccountName    = $SamAccountName
        UserPrincipalName = $Email
        EmailAddress      = $Email
        Description       = $Class
        Company           = "Student"
        EmployeeID        = "$StudentID"
        ScriptPath        = "student_login_script.bat"
        HomeDirectory     = $HomeDirectory
        HomeDrive         = "X:"
        Path              = $ouPath
        AccountPassword   = (ConvertTo-SecureString -AsPlainText $Password -Force)
        Enabled           = $true
    }

    try {
        New-ADUser @userParams -Server $env:USERDNSDOMAIN

        Set-ADUser -Identity $SamAccountName -Replace @{
            Pager        = $StudentID
            MailNickname = $SamAccountName
        }

        $user = Get-ADUser -Identity $SamAccountName -Properties UserAccountControl
        $userAccountControl = $user.UserAccountControl -bor 0x10000
        Set-ADUser -Identity $SamAccountName -Replace @{userAccountControl = $userAccountControl}

        # Placeholder group names
        $groups = @(
            "PRINTING-STUDENTS", "Student_NetworkAccess",
            "Office365_Students", "STUDENT_WIFI", "STUDENTS_GROUP", "VPN_Students"
        )
        foreach ($group in $groups) {
            Add-ADGroupMember -Identity $group -Members $SamAccountName
        }
        Write-Log "User $SamAccountName added successfully as $Email"
    }
    catch {
        $shortError = $_.Exception.Message
        $fullError = $_ | Out-String

        Write-Host ""
        Write-Host "❌ Error adding user: $SamAccountName" -ForegroundColor Red
        Write-Host "Message: $shortError"
        Write-Host "----- FULL ERROR DETAILS -----"
        Write-Host $fullError
        Write-Host "----- INPUT VALUES -----"
        $userParams.GetEnumerator() | ForEach-Object { Write-Host "$($_.Key): $($_.Value)" }
        Write-Host "------------------------------"
    }
}

# Import and process users
try {
    $users = Import-Csv -Path $CsvPath -Header "Password", "FirstName", "LastName", "StudentID", "Initial", "Class" | Select-Object -Skip 1

    foreach ($user in $users) {
        if ($user.Password -and $user.FirstName -and $user.LastName) {
            $UserArray = @(
                $user.Password,
                $user.FirstName,
                $user.LastName,
                $user.StudentID,
                $user.Initial,
                $user.Class
            )

            Add-Mass-ADStudentSafely -UserData $UserArray
        } else {
            Write-Host "⚠️ Skipped incomplete row." -ForegroundColor Yellow
        }
    }

    Write-Host "All users processed." -ForegroundColor Green
    return "SUCCESS"
}
catch {
    Write-Host "❌ Failed to process import: $($_ | Out-String)" -ForegroundColor Red
    exit 1
}
