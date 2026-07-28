<#
    Syncs the First Pass Q Report straight from SharePoint Online into Supabase.

    Reads the workbook over Microsoft Graph - no local copy, no OneDrive sync, no
    downloaded file. Uses the Microsoft Graph PowerShell module's own Microsoft-owned
    app registration, so nothing needs registering in Azure. You sign in once, in
    Microsoft's prompt; the token is cached for later unattended runs.

    One-time setup:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
        .\sync-first-pass-online.ps1        # sign in when prompted

    Then schedule with register-sync-task.ps1 (point $Script at this file).

    Each run first asks Graph when the workbook last changed and exits early if it
    hasn't - so polling every couple of minutes costs one tiny API call.
#>

# --- Settings ---------------------------------------------------------------
$SiteHost     = 'mycopeland.sharepoint.com'
$SitePath     = '/sites/PanelServices'
$LibraryName  = 'Documents'                                  # "Shared Documents" in URLs
$FilePath     = 'General/Copeland First Pass Q Report.xlsx'   # within the library
$SheetData    = 'Daily Results'
$SheetNames   = 'Names'

$SupabaseUrl  = 'https://ptbhguthosenkffjhbry.supabase.co'
$AnonKey      = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InB0YmhndXRob3NlbmtmZmpoYnJ5Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ5MDkzNzAsImV4cCI6MjEwMDQ4NTM3MH0.y71yXWsIUjzFBqC4itzZ36w9ixw_QDej2RaBPh8MLPI'
# Loaded from sync-config.local.ps1, which is gitignored. This repo is public,
# so the secret must never be written into a tracked file.
$IngestSecret = $null

$BatchSize    = 5000
$LogPath      = Join-Path $PSScriptRoot 'sync-first-pass.log'
$StampPath    = Join-Path $PSScriptRoot '.last-modified'
$ForceSync    = $false     # set $true to sync even when unchanged

# --- Helpers -----------------------------------------------------------------
$ErrorActionPreference = 'Stop'

# Pull the secret from the machine-local, untracked config.
$cfg = Join-Path $PSScriptRoot 'sync-config.local.ps1'
if (Test-Path $cfg) { . $cfg }
if (-not $IngestSecret) {
    throw "No ingest secret. Create $cfg containing:  `$IngestSecret = '<secret>'"
}

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
        'Authorization' = "Bearer $AnonKey"
        'Content-Type'  = 'application/json'
    }
    Invoke-RestMethod -Method Post -Uri "$SupabaseUrl/rest/v1/rpc/$Function" `
                      -Headers $headers -Body ($Body | ConvertTo-Json -Depth 6 -Compress)
}

function ConvertTo-IsoDate {
    param($Value)
    if ($null -eq $Value -or "$Value".Trim() -eq '') { return $null }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd') }
    $n = 0.0
    if ([double]::TryParse("$Value", [ref]$n) -and $n -gt 0) {
        return [datetime]::FromOADate($n).ToString('yyyy-MM-dd')   # Excel serial
    }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse("$Value", [ref]$d)) { return $d.ToString('yyyy-MM-dd') }
    return $null
}

function AsText { param($v) if ($null -eq $v) { '' } else { "$v".Trim() } }

# Reads a whole worksheet as a 2D array without downloading the file.
function Get-UsedRange {
    param([string]$ItemUri, [string]$Sheet)
    $uri = "$ItemUri/workbook/worksheets('$([uri]::EscapeDataString($Sheet))')/usedRange(valuesOnly=true)?`$select=values"
    (Invoke-MgGraphRequest -Method GET -Uri $uri).values
}

# --- Run ---------------------------------------------------------------------
try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Missing module. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force'
    }
    Import-Module Microsoft.Graph.Authentication

    # Reuses the cached token when there is one; prompts in a browser otherwise.
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes 'Files.Read.All','Sites.Read.All' -NoWelcome
    }
    Write-Log ("Signed in as {0}" -f (Get-MgContext).Account)

    # Resolve site -> library -> file
    $site  = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/${SiteHost}:${SitePath}"
    $drives = (Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$($site.id)/drives").value
    $drive = $drives | Where-Object { $_.name -eq $LibraryName } | Select-Object -First 1
    if (-not $drive) {
        throw ("Library '$LibraryName' not found. Available: " + (($drives.name) -join ', '))
    }

    $encoded  = ($FilePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    $itemUri  = "/v1.0/drives/$($drive.id)/root:/${encoded}:"
    $item     = Invoke-MgGraphRequest -Method GET -Uri "${itemUri}?`$select=id,name,lastModifiedDateTime,size"
    $modified = $item.lastModifiedDateTime

    # Skip the heavy work when nothing changed - makes frequent polling cheap.
    $previous = if (Test-Path $StampPath) { (Get-Content $StampPath -Raw).Trim() } else { '' }
    if (-not $ForceSync -and $previous -eq "$modified") {
        Write-Log "Unchanged since $previous - nothing to do."
        return
    }
    Write-Log ("Workbook changed ({0}), reading..." -f $modified)

    $itemUri = "/v1.0/drives/$($drive.id)/items/$($item.id)"
    $values  = Get-UsedRange -ItemUri $itemUri -Sheet $SheetData
    if (-not $values -or $values.Count -lt 2) { throw "No rows returned from '$SheetData'." }

    # Read by column position, matching the dashboard's own parser.
    $rows = for ($i = 1; $i -lt $values.Count; $i++) {
        $r   = $values[$i]
        $iso = ConvertTo-IsoDate $r[1]
        if (-not $iso) { continue }
        [ordered]@{
            panel = AsText $r[0];  date          = $iso
            pass_fail = AsText $r[2];  reason    = AsText $r[3]
            line = AsText $r[4];   emp           = AsText $r[5]
            routing_hours = AsText $r[6];  tech  = AsText $r[7]
            layout = AsText $r[8]; wire          = AsText $r[9]
            fa = AsText $r[10];    sales_order   = AsText $r[11]
        }
    }
    $rows = @($rows)
    if ($rows.Count -eq 0) { throw 'No dated rows found - check the Date column.' }

    # Names sheet: locate Initials/Name by header text.
    $names = @()
    try {
        $nv = Get-UsedRange -ItemUri $itemUri -Sheet $SheetNames
        if ($nv -and $nv.Count -gt 1) {
            $hdr = @($nv[0] | ForEach-Object { "$_".Trim().ToLower() })
            $ci  = $hdr.IndexOf('initials'); $cn = $hdr.IndexOf('name')
            if ($ci -ge 0 -and $cn -ge 0) {
                for ($i = 1; $i -lt $nv.Count; $i++) {
                    $ini = AsText $nv[$i][$ci]; $nm = AsText $nv[$i][$cn]
                    if ($ini -and $nm) { $names += [ordered]@{ initials = $ini; name = $nm } }
                }
            }
        }
    } catch { Write-Log "Names sheet not read: $($_.Exception.Message)" 'WARN' }

    Write-Log ("Parsed {0} rows, {1} names" -f $rows.Count, $names.Count)

    # Staged load: live data is replaced only once the whole run succeeds.
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

    Set-Content -Path $StampPath -Value "$modified" -Encoding ascii
    Write-Log ("Published {0} rows (was {1}), {2} names" -f `
               $final.rows, $final.prev_rows, $final.names) 'OK'
}
catch {
    Write-Log $_.Exception.Message 'ERROR'
    exit 1
}
