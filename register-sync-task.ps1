<#
    Registers sync-first-pass.ps1 to run every 15 minutes with no visible window.

    Launches through `conhost.exe --headless`, which runs the command without
    creating a console at all - so nothing appears or steals focus while you work.
    Falls back to `-WindowStyle Hidden` if conhost is unavailable (that can still
    flash briefly on some builds).

    Per-user task: no admin rights needed. Runs while you are logged in.

    Run once:   .\register-sync-task.ps1
    Remove:     Unregister-ScheduledTask -TaskName 'First Pass Sync' -Confirm:$false
#>

$ErrorActionPreference = 'Stop'

$TaskName    = 'First Pass Sync'
$IntervalMin = 15
$Script      = Join-Path $PSScriptRoot 'sync-first-pass.ps1'
$LogPath     = Join-Path $PSScriptRoot 'sync-first-pass.log'

if (-not (Test-Path $Script)) { throw "Not found: $Script" }

$psArgs  = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}"' -f $Script
$conhost = Join-Path $env:SystemRoot 'System32\conhost.exe'

if (Test-Path $conhost) {
    $exec = $conhost
    $arg  = '--headless powershell.exe ' + $psArgs
    $mode = 'conhost --headless (no console window is created)'
} else {
    $exec = 'powershell.exe'
    $arg  = '-WindowStyle Hidden ' + $psArgs
    $mode = 'hidden window (a console may flash briefly)'
}

$action  = New-ScheduledTaskAction -Execute $exec -Argument $arg
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
    -RepetitionInterval (New-TimeSpan -Minutes $IntervalMin)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
    -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Description 'Syncs the First Pass Q Report into Supabase.' `
    -Force | Out-Null

# Report what is actually registered, not what we intended - a failed
# registration can otherwise leave an older task in place unnoticed.
$t = Get-ScheduledTask -TaskName $TaskName
$ok = $t.Actions[0].Execute -eq $exec
Write-Host ""
Write-Host "Registered '$TaskName' - every $IntervalMin minutes." -ForegroundColor Green
Write-Host "  Intended  : $mode"
Write-Host "  Executes  : $($t.Actions[0].Execute)"
Write-Host "  Arguments : $($t.Actions[0].Arguments)"
Write-Host "  Logon type: $($t.Principal.LogonType)"
if (-not $ok) {
    Write-Host "  WARNING: registered action does not match - an older task may still be in place." -ForegroundColor Yellow
}
Write-Host ""
Write-Host "Run it now:  Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "Log file:    $LogPath"
