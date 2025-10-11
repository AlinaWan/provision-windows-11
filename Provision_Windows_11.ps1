<#
    Windows 11 System Info Checker and Provisoning Script
    -----------------------------------------------------

    Desired State Definitions:

    1. Vivaldi Installation:
       - Vivaldi should be installed for the current user.
       - If not installed, it will be reported as [WARN].

    2. Default Applications:
       - Vivaldi should be set as the default application for the following file extensions and protocols:
         - .htm
         - .html
         - .pdf
         - .svg
         - .webp
         - http
         - https
       - If any of these are not set to Vivaldi, it will be reported as [WARN].
       - ProgID values are always displayed for reference.

    3. Taskbar Alignment:
       - Taskbar should be aligned to the LEFT.
       - Any other alignment is reported as [WARN].
       - TaskbarAl value is displayed.

    4. Dark Mode:
       - Dark mode should be enabled (AppsUseLightTheme = 0).
       - If dark mode is not enabled, it will be reported as [WARN].
       - AppsUseLightTheme value is displayed.

    4. Clipboard History:
       - Clipboard history should be enabled (EnableClipboardHistory = 1).
       - If clipboard history is not enabled, it will be reported as [WARN].
       - EnableClipboardHistory value is displayed.

    5. File Explorer Settings:
       - Hidden files and folders should be visible (Hidden = 1).
       - File extensions for known file types should be displayed (HideFileExt = 0).
       - If any of these settings do not match the desired state, it will be reported as [WARN].
       - Current registry values for each setting are displayed for reference.

    Reporting Levels:
       - [OK]    : Matches the desired state exactly.
       - [WARN]  : Valid state but not the desired configuration.
       - [ERROR] : Unable to determine the setting or read the system value.

    Notes:
       - Remediation is attempted if any value does not match the desired state.
       - All technical values are printed prior to the determination.
       - The script checks only the current user scope for Vivaldi and default apps.
#>


function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    switch ($Level.ToUpper()) {
        "INFO" { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
        "WARN" { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
        "ERROR" { Write-Host "[ERROR] $Message" -ForegroundColor Red }
        "OK" { Write-Host "[OK] $Message" -ForegroundColor Green }
        "SUCCESS" { Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
        default { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
    }
}

# Desired settings
$DesiredDefaults = @{
    ".htm"  = "Vivaldi"
    ".html" = "Vivaldi"
    ".pdf"  = "Vivaldi"
    ".svg"  = "Vivaldi"
    ".webp" = "Vivaldi"
    "http"  = "Vivaldi"
    "https" = "Vivaldi"
}
$DesiredTaskbar = 0 # 0 = LEFT
$DesiredDarkMode = 0 # 0 = dark mode on
$DesiredClipboard = 1 # 1 = enabled

# --- Checking Vivaldi installation (user scope) ---
Write-Host "`n--- Checking Vivaldi installation (user scope) ---" -ForegroundColor Yellow
$vivaldiRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
$vivaldiKey = Get-ChildItem $vivaldiRegPath | Where-Object {
    (Get-ItemProperty $_.PSPath).DisplayName -like "Vivaldi*"
}
if ($vivaldiKey) {
    $installPath = (Get-ItemProperty $vivaldiKey.PSPath).InstallLocation
    Write-Host "InstallLocation = $installPath"
    Write-Log "Vivaldi is installed." "OK"
    $vivaldiInstalled = $true
} else {
    Write-Host "Vivaldi registry key not found."
    Write-Log "Vivaldi is not installed." "WARN"
    Write-Log "Installing Vivaldi for current user..." "INFO"
    try {
        winget install Vivaldi.Vivaldi --scope user -e --silent
        Write-Log "Vivaldi installation attempted." "SUCCESS"
        $vivaldiInstalled = $true
    } catch {
        Write-Log "Failed to install Vivaldi." "ERROR"
        $vivaldiInstalled = $false
    }
}

# --- Default apps ---
function Check-DefaultApp {
    param (
        [string]$item,
        [string]$desired
    )

    try {
        if ($item -match "http|https") {
            $regPath = "HKCU:\Software\Microsoft\Windows\Shell\Associations\UrlAssociations\$item\UserChoice"
            $progId = (Get-ItemProperty -Path $regPath -ErrorAction Stop).ProgId
        } else {
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$item\UserChoice"
            $progId = (Get-ItemProperty -Path $regPath -ErrorAction Stop).ProgId
        }

        Write-Host "ProgID = $progId"

        if ($progId -like "$desired*") {
            Write-Log "$item is set to Vivaldi (default)." "OK"
        } else {
            Write-Log "$item is NOT set to Vivaldi." "WARN"
            Write-Log "Attempting to set $item to Vivaldi..." "INFO"
            try {
                # Windows 11 user defaults cannot be easily set programmatically
                Start-Process "ms-settings:defaultapps" -Wait
                Write-Log "$item default change triggered." "SUCCESS"
            } catch {
                Write-Log "Failed to set $item default." "ERROR"
            }
        }
    } catch {
        Write-Host "Unable to read ProgID"
        Write-Log "Could not determine default app for $item." "ERROR"
    }
}

if ($vivaldiInstalled) {
    Write-Host "`n--- Checking default apps for Vivaldi ---" -ForegroundColor Yellow
    foreach ($item in $DesiredDefaults.Keys) {
        Check-DefaultApp -item $item -desired $DesiredDefaults[$item]
    }
}

# --- Taskbar alignment ---
Write-Host "`n--- Checking taskbar alignment ---" -ForegroundColor Yellow
try {
    $alignment = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarAl -ErrorAction Stop).TaskbarAl
    Write-Host "TaskbarAl = $alignment"
    if ($alignment -eq $DesiredTaskbar) {
        Write-Log "Taskbar is aligned LEFT." "OK"
    } else {
        Write-Log "Taskbar alignment is not LEFT." "WARN"
        Write-Log "Attempting to set taskbar alignment to LEFT..." "INFO"
        try {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name TaskbarAl -Value $DesiredTaskbar
            Stop-Process -ProcessName explorer -Force
            Write-Log "Taskbar alignment set to LEFT." "SUCCESS"
        } catch {
            Write-Log "Failed to set taskbar alignment." "ERROR"
        }
    }
} catch {
    Write-Host "TaskbarAl not found"
    Write-Log "Taskbar alignment setting not found." "ERROR"
}

# --- Dark mode ---
Write-Host "`n--- Checking dark mode status ---" -ForegroundColor Yellow
try {
    $appsUseLightTheme = Get-ItemProperty "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -ErrorAction Stop
    Write-Host "AppsUseLightTheme = $($appsUseLightTheme.AppsUseLightTheme)"
    if ($appsUseLightTheme.AppsUseLightTheme -eq $DesiredDarkMode) {
        Write-Log "Dark mode is enabled." "OK"
    } else {
        Write-Log "Dark mode is not enabled." "WARN"
        Write-Log "Enabling dark mode..." "INFO"
        try {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value $DesiredDarkMode
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value $DesiredDarkMode
            Write-Log "Dark mode enabled." "SUCCESS"
        } catch {
            Write-Log "Failed to enable dark mode." "ERROR"
        }
    }
} catch {
    Write-Host "AppsUseLightTheme not found"
    Write-Log "Unable to determine dark mode status." "ERROR"
}

# --- Clipboard history ---
Write-Host "`n--- Checking clipboard history ---" -ForegroundColor Yellow
try {
    $clipboardValue = (Get-ItemProperty "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -ErrorAction Stop).EnableClipboardHistory
    Write-Host "EnableClipboardHistory = $clipboardValue"
    if ($clipboardValue -eq $DesiredClipboard) {
        Write-Log "Clipboard history is enabled." "OK"
    } else {
        Write-Log "Clipboard history is not enabled." "WARN"
        Write-Log "Enabling clipboard history..." "INFO"
        try {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Clipboard" -Name "EnableClipboardHistory" -Value $DesiredClipboard
            Write-Log "Clipboard history enabled." "SUCCESS"
        } catch {
            Write-Log "Failed to enable clipboard history." "ERROR"
        }
    }
} catch {
    Write-Host "EnableClipboardHistory not found"
    Write-Log "Unable to determine clipboard history status." "ERROR"
}

# --- File Explorer settings ---
Write-Host "`n--- Checking File Explorer settings ---" -ForegroundColor Yellow
$explorerRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"

# Desired values
$desiredSettings = @{
    "Hidden"          = 1  # Show hidden files and folders
    "HideFileExt"     = 0  # Show file extensions for known types
}

$changesMade = $false  # Flag to track if any changes happen

foreach ($setting in $desiredSettings.Keys) {
    try {
        $currentValue = (Get-ItemProperty -Path $explorerRegPath -Name $setting -ErrorAction Stop).$setting
        Write-Host "$setting = $currentValue"
        if ($currentValue -ne $desiredSettings[$setting]) {
            Write-Log "$setting is not as desired." "WARN"
            Write-Log "Setting $setting to $($desiredSettings[$setting])..." "INFO"
            Set-ItemProperty -Path $explorerRegPath -Name $setting -Value $desiredSettings[$setting]
            Write-Log "$setting updated successfully." "SUCCESS"
            $changesMade = $true
        } else {
            Write-Log "$setting is correctly configured." "OK"
        }
    } catch {
        Write-Log "Failed to read or set $setting." "ERROR"
    }
}

# Refresh Explorer only if any changes were made
if ($changesMade) {
    Write-Log "Changes were made. Restarting Explorer to apply settings..." "INFO"
    Stop-Process -ProcessName explorer -Force
} else {
    Write-Log "No changes made. Explorer restart not required." "INFO"
}