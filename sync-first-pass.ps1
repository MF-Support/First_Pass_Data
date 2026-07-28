<#
    Syncs the First Pass Q Report workbook into Supabase.

    Replaces the Power Automate flow entirely - no premium licence, no Azure app
    registration, no stored password. OneDrive already keeps the workbook current
    on this machine; this reads that local copy and pushes it up.

    One-time setup:
        Install-Module ImportExcel -Scope CurrentUser -Force
    Then set $Workbook below and run. Schedule it with register-sync-task.ps1.
#>

# --- Settings ---------------------------------------------------------------
# Path to the OneDrive-synced workbook. Find it in Explorer, Shift+Right-click
# the file, "Copy as path", and paste between the quotes.
$Workbook    = "$env:USERPROFILE\Copeland\Panel Services - General\Copeland First Pass Q Report.xlsx"

$SupabaseUrl = 'https://ptbhguthosenkffjhbry.supabase.co'
$AnonKey     = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB0YmhndXRob3NlbmtmZmpoYnJ5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ5MDkzNzAsImV4cCI6MjEwMDQ4NTM3MH0.y71yXWsIUjzFBqC4itzZ36w9ixw_QDej2RaBPh8MLPI'
$IngestSecret= '2769cca6ec7318717d562a43548d1428a3babe3d85ced4cf'

$SheetData   = 'Daily Results'
$SheetNames  = 'Names'
$BatchSize   = 5000
$LogPath     = Join-Path $PSScriptRoot 'sync-first-pass.log'

# --- Helpers -----------------------------------------------------------------
$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -Path $LogPath -Value $line -Encoding utf8
}

function Invoke-Rpc {
    param([string]$Function, [hashtable]$Body)
    $headers = @{
        'apikey'        = $AnonKey
        'Authorization' = "Bearer $AnonKey"   # one space, not a plus
        'Content-Type'  = 'application/json'
    }
    $json = $Body | ConvertTo-Json -Depth 6 -Compress
    return Invoke-RestMethod -Method Post -Uri "$SupabaseUrl/rest/v1/rpc/$Function" `
                             -Headers $headers -Body $json
}

# Excel serial / DateTime / text -> yyyy-MM-dd, or $null when unreadable.
function ConvertTo-IsoDate {
    param($Value)
    if ($null -eq $Value -or "$Value".Trim() -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $n = 0.0
    if ([double]::TryParse("$Value", [ref]$n) -and $n -gt 0) {
        return [datetime]::FromOADate($n).ToString('yyyy-MM-dd')
    }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse("$Value", [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
    return $null
}

function AsText { param($v) if ($null -eq $v) { '' } else { "$v".Trim() } }

# --- Run ---------------------------------------------------------------------
try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw 'ImportExcel module missing. Run: Install-Module ImportExcel -Scope CurrentUser -Force'
    }
    Import-Module ImportExcel

    if (-not (Test-Path $Workbook)) { throw "Workbook not found: $Workbook" }

    # Work from a copy: the live file is usually open in Excel or mid-sync.
    $temp = Join-Path $env:TEMP ("fp-sync-{0}.xlsx" -f (Get-Date -Format 'yyyyMMddHHmmss'))
    Copy-Item -LiteralPath $Workbook -Destination $temp -Force
    Write-Log "Read $Workbook"

    try {
        # Read by column position, matching the dashboard's own parser, so a
        # renamed header can't silently change what gets synced.
        $raw = Import-Excel -Path $temp -WorksheetName $SheetData -NoHeader -StartRow 2

        $rows = foreach ($r in $raw) {
            $iso = ConvertTo-IsoDate $r.P2
            if (-not $iso) { continue }          # blank/junk date -> skip, as the dashboard does
            [ordered]@{
                panel         = AsText $r.P1
                date          = $iso
                pass_fail     = AsText $r.P3
                reason        = AsText $r.P4
                line          = AsText $r.P5
                emp           = AsText $r.P6
                routing_hours = AsText $r.P7
                tech          = AsText $r.P8
                layout        = AsText $r.P9
                wire          = AsText $r.P10
                fa            = AsText $r.P11
                sales_order   = AsText $r.P12
            }
        }
        $rows = @($rows)
        if ($rows.Count -eq 0) { throw 'No dated rows found - check $SheetData and the Date column.' }

        # Names sheet: match the Initials/Name columns by header.
        $names = @()
        try {
            foreach ($n in (Import-Excel -Path $temp -WorksheetName $SheetNames)) {
                $ini = AsText $n.Initials
                $nm  = AsText $n.Name
                if ($ini -and $nm) { $names += [ordered]@{ initials = $ini; name = $nm } }
            }
        } catch { Write-Log "Names sheet not read: $($_.Exception.Message)" 'WARN' }

        Write-Log ("Parsed {0} rows, {1} names" -f $rows.Count, $names.Count)

        # Staged load: live data is untouched until commit, so a failure part way
        # through leaves the previous sync intact rather than a half-loaded table.
        $token = (Invoke-Rpc 'fp_ingest_begin' @{ p_secret = $IngestSecret }).token

        for ($i = 0; $i -lt $rows.Count; $i += $BatchSize) {
            $batch = $rows[$i..([Math]::Min($i + $BatchSize - 1, $rows.Count - 1))]
            $res = Invoke-Rpc 'fp_ingest_chunk' @{
                p_secret = $IngestSecret; p_token = $token; p_rows = @($batch)
            }
            Write-Log ("  staged {0}/{1}" -f $res.staged_total, $rows.Count)
        }

        $final = Invoke-Rpc 'fp_ingest_commit' @{
            p_secret = $IngestSecret; p_token = $token; p_names = @($names)
        }
        Write-Log ("Published {0} rows (was {1}), {2} names" -f `
                   $final.rows, $final.prev_rows, $final.names) 'OK'
    }
    finally {
        Remove-Item $temp -Force -ErrorAction SilentlyContinue
    }
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    exit 1
}
