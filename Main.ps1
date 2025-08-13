# **********************************************************************
# *                         Windows API Setup                        *
# **********************************************************************

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    public const int SW_HIDE = 0;
    public const int SW_SHOW = 5;
}
"@

# **********************************************************************
# *                     Console Window Initialization                 *
# **********************************************************************

# Store handle as a global variable
$global:consoleHandle = [Win32]::GetConsoleWindow()

# Initially hide the console window, suppress output
[Win32]::ShowWindow($global:consoleHandle, [Win32]::SW_HIDE) | Out-Null

# **********************************************************************
# *                        Console Warning Banner                     *
# **********************************************************************

$bgColor = "DarkRed"
$fgColor = "White"

$bannerWidth = 99
$bannerLine = "*" * $bannerWidth
$message = "* DO NOT CLOSE THIS WINDOW - this will stop the application *"

Write-Host $bannerLine -BackgroundColor $bgColor -ForegroundColor $fgColor
Write-Host $message -BackgroundColor $bgColor -ForegroundColor $fgColor
Write-Host $bannerLine -BackgroundColor $bgColor -ForegroundColor $fgColor
Write-Host ""  # Blank line for spacing

# **********************************************************************
# *                            Module Imports                         *
# **********************************************************************

. "$PSScriptRoot\UI_Module.ps1"
. "$PSScriptRoot\AD_Helper.ps1"
. "$PSScriptRoot\Logger_Module.ps1"

# **********************************************************************
# *                          Launch User Form                         *
# **********************************************************************

$exitCode = Show-UserForm

if ($exitCode -eq "Closed") {
    Write-Log "User form closed. Exiting script."
    Write-Host "User form closed. Exiting script."
    exit 0
}
elseif ($exitCode -eq "Success") {
    Write-Log "User import completed successfully."
    Write-Host "Import finished successfully. Exiting script."
    exit 0
}
else {
    Write-Log "Unexpected exit code: $exitCode"
    Write-Host "⚠️ Unexpected result: $exitCode"
    exit 1
}
