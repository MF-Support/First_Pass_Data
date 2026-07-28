<#
    Registers sync-first-pass.ps1 to run every 15 minutes.

    Creates a per-user scheduled task - no admin rights needed. Runs only while
    you are logged in, which suits a shop-floor PC that stays signed in.

    Run once:   .\register-sync-task.ps1
    Remove:     Unregister-ScheduledTask -TaskName 'First Pass Sync' -Confirm:$false
#>

$ErrorActionPreference = 'Stop'

$TaskName    = 'First Pass Sync'
$IntervalMin = 15
$Script      = Join-Path $PSScriptRoot 'sync-first-pass.ps1'
$LogPath     = Join-Path $PSScriptRoot 'sync-first-pass.log'

if (-not (Test-Path $Script)) { throw "Not found: $Script" }

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument ('-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $Script)

# Repeat indefinitely, starting a minute from now.
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMin)

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Description 'Syncs the First Pass Q Report into Supabase.' `
    -Force | Out-Null

Write-Host "Registered '$TaskName' - runs every $IntervalMin minutes." -ForegroundColor Green
Write-Host "Run it now:  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Log file:    $LogPath"
