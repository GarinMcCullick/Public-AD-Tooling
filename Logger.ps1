# Create log folder
$LogFolder = "$PSScriptRoot\Logs"
if (!(Test-Path -Path $LogFolder)) {
    New-Item -ItemType Directory -Path $LogFolder -Force
}

# Get the current PowerShell session user
$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if (-not $Username) {
    $Username = $env:USERNAME  # Fallback
}
$Username = ($Username -split '\\')[-1]  # Remove domain if present

# Get current date and time in a chosen time zone
$TimeZoneId = "Your Time Zone Name"  # Example: "Pacific Standard Time"
$CurrentTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, $TimeZoneId)

# Log file path includes username
$LogFileName = "{0}_{1}_{2}.log" -f $CurrentTime.ToString("MM-dd-yyyy"), $Username, $CurrentTime.ToString("h-mm-ss_tt")
$LogFilePath = "$LogFolder\$LogFileName"

function Write-Log {
    param (
        [string]$Message,
        [string]$LogType = "Info"
    )

    # Refresh current time for each log entry
    $CurrentTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId([DateTime]::UtcNow, $TimeZoneId)
    $TimeStamp = $CurrentTime.ToString("MM-dd-yyyy h:mm:ss tt")
    $LogEntry = "$TimeStamp [$LogType] $Message"

    try {
        Add-Content -Path $LogFilePath -Value $LogEntry -ErrorAction Stop
    } catch {
        Write-Host "Failed to write to log file: $LogFilePath" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Color-coded console output
    switch ($LogType) {
        "Info"    { Write-Host $LogEntry -ForegroundColor Green }
        "Warning" { Write-Host $LogEntry -ForegroundColor Yellow }
        "Error"   { Write-Host $LogEntry -ForegroundColor Red }
        default   { Write-Host $LogEntry }
    }
}
