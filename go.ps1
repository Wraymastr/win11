# ==============================
# Admin Elevation Check
# ==============================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Please run this script as Administrator!"
    break
}

# ==============================
# Winget reliability tweak
# ==============================
Write-Host "[Winget] Enabling bypasscertificatepinningformicrosoftstore..." -ForegroundColor Cyan
winget settings --enable bypasscertificatepinningformicrosoftstore
Start-Sleep -Seconds 2

# ==============================
# Privacy, Telemetry & Ads
# ==============================
Write-Host "[Privacy] Disabling Telemetry, Ads, and Tailored Experiences..." -ForegroundColor Cyan
# Telemetry
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" -Name "AllowTelemetry" -Value 0 -ErrorAction SilentlyContinue
Set-Service -Name "DiagTrack" -StartupType Disabled -ErrorAction SilentlyContinue

# Ad ID and Tailored Experiences
$AdsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
Set-ItemProperty -Path $AdsPath -Name "SubscribedContent-310093Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $AdsPath -Name "SubscribedContent-353696Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0 -ErrorAction SilentlyContinue

# ==============================
# Interface Cleanup (Taskbar & Lock Screen)
# ==============================
Write-Host "[UI] Cleaning Taskbar and Lock Screen..." -ForegroundColor Cyan
# Disable Widgets (News and Interests)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds" -Name "ShellFeedsTaskbarViewMode" -Value 2 -ErrorAction SilentlyContinue

# Remove Lock Screen "Facts and Tips"
Set-ItemProperty -Path $AdsPath -Name "RotatingLockScreenOverlayEnabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $AdsPath -Name "SubscribedContent-338387Enabled" -Value 0 -ErrorAction SilentlyContinue

# Hide Meet Now and People
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "PeopleBand" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "HideSCAMeetNow" -Value 1 -ErrorAction SilentlyContinue

#Declutter the Search Bar
Write-Host "[UI] Disabling Search Highlights and Web Search..." -ForegroundColor Cyan
# Disable Search Highlights (the icons in the search bar)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings" -Name "IsDynamicSearchBoxPresent" -Value 0
# Disable Bing Search in the Start Menu (keeps search local and fast)
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "BingSearchEnabled" -Value 0

# ==============================
# Performance Tweaks
# ==============================
Write-Host "[Perf] Disabling Hibernation and Edge Startup Boost..." -ForegroundColor Cyan
powercfg.exe /hibernate off
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "StartupBoostEnabled" -Value 0 -ErrorAction SilentlyContinue

Write-Host "[Config] Setting Active Hours (8 AM - 5 PM)..." -ForegroundColor Cyan
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "ActiveHoursStart" -Value 8
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "ActiveHoursEnd" -Value 17

# ==============================
# Remove Microsoft bloat (VERBOSE)
# ==============================
function Remove-MicrosoftBloat {
    Write-Host "[Bloat] Starting Microsoft bloat removal..." -ForegroundColor Yellow

    $bloatApps = @(
        "*MicrosoftTeams*", "*MSTeams*", "*Clipchamp*", "*Xbox*", "*Cortana*", 
        "*MicrosoftNews*", "*BingWeather*", "*GetHelp*", "*FeedbackHub*", 
        "*WindowsMaps*", "*YourPhone*", "*MicrosoftToDo*", "*MicrosoftOfficeHub*", 
        "*People*", "*ZuneMusic*", "*ZuneVideo*", "Microsoft.Messaging", "Microsoft.OneConnect"
    )

    foreach ($app in $bloatApps) {
        Write-Host "[Bloat] Processing $app..."
        $packages = Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue
        if ($packages) {
            foreach ($pkg in $packages) {
                Write-Host "    Removing installed package: $($pkg.Name)"
                Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction SilentlyContinue
            }
        }

        $provPkgs = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like $app }
        if ($provPkgs) {
            foreach ($prov in $provPkgs) {
                Write-Host "    Removing provisioned package: $($prov.DisplayName)"
                Remove-AppxProvisionedPackage -Online -PackageName $prov.PackageName -ErrorAction SilentlyContinue
            }
        }
    }
}

Remove-MicrosoftBloat

# ==============================
# Uninstall OneDrive
# ==============================
Write-Host "[OneDrive] Uninstalling OneDrive..." -ForegroundColor Yellow
$oneDriveSetup = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
if (-not (Test-Path $oneDriveSetup)) { $oneDriveSetup = "$env:SystemRoot\System32\OneDriveSetup.exe" }

if (Test-Path $oneDriveSetup) {
    Start-Process $oneDriveSetup "/uninstall" -NoNewWindow -Wait
    Write-Host "[OneDrive] Uninstall complete"
}
# ==============================
# Autostart Cleanup (Legacy Run Keys)
# ==============================
Write-Host "[Autostart] Removing leftover run keys..." -ForegroundColor Cyan
$RunPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
Remove-ItemProperty -Path $RunPath -Name "OneDrive" -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $RunPath -Name "com.squirrel.Teams.Teams" -ErrorAction SilentlyContinue

# ==============================
# HP Specific Bloatware Removal
# ==============================
Write-Host "[HP-Cleanup] Neutralizing HP Wolf Security and Bloat..." -ForegroundColor Yellow

$hpApps = @("HP.HPSupportAssistant", "HP.HPQuickDrop", "HP.HPPrivacySettings", "HP.HPConnectionOptimizer", "HP.HPSmart")

foreach ($app in $hpApps) {
    Write-Host "  Processing $app..."
    # Removed --ignore-uninstalled for better compatibility
    winget uninstall --id $app -e --accept-source-agreements -h --force
}

# Targeted WMI cleanup for the stubborn Wolf Security
$wolf = Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -like "*HP Wolf Security*" -or $_.Name -like "*HP Sure Sense*" }
if ($wolf) {
    foreach ($item in $wolf) {
        Write-Host "    Force uninstalling $($item.Name)..."
        Invoke-CimMethod -InputObject $item -MethodName "Uninstall" | Out-Null
    }
}

# ==============================
# Install apps
# ==============================
Write-Host "[Install] Updating/Installing Core Apps..." -ForegroundColor Green

# Use a simple array to loop through installs
$appsToInstall = @(
    "Google.Chrome",
    "Google.GoogleDrive",
    "VideoLAN.VLC",
    "Mozilla.Firefox"
)

foreach ($id in $appsToInstall) {
    Write-Host "  Installing $id..."
    winget install --id $id -e --accept-source-agreements --accept-package-agreements
}

Write-Host "All tasks complete! Please restart to see the full effect. ??" -ForegroundColor Magenta
