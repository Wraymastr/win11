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

# ==============================
# Performance Tweaks
# ==============================
Write-Host "[Perf] Disabling Hibernation and Edge Startup Boost..." -ForegroundColor Cyan
powercfg.exe /hibernate off
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "StartupBoostEnabled" -Value 0 -ErrorAction SilentlyContinue

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
# Install apps
# ==============================
Write-Host "[Install] Installing Google Chrome..." -ForegroundColor Green
winget install --id Google.Chrome -e --accept-source-agreements --accept-package-agreements

Write-Host "[Install] Installing Google Drive for Desktop..." -ForegroundColor Green
winget install --id Google.GoogleDrive -e --accept-source-agreements --accept-package-agreements

Write-Host "[Install] Installing Firefox..." -ForegroundColor Green
winget install --id Mozilla.Firefox -e --accept-source-agreements --accept-package-agreements

Write-Host "[Install] Installing VLC..." -ForegroundColor Green
winget install --id VideoLAN.vlc -e --accept-source-agreements --accept-package-agreements

Write-Host "All tasks complete! Please restart to see the full effect. ??" -ForegroundColor Magenta
