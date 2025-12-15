<#
.SYNOPSIS
    Import script that updates AD user signature/profile attributes from a CSV exported by Export-ADUsers-SignatureData.ps1
    (typically edited in Excel) for use with email signature solutions (e.g., CodeTwo Signatures and similar tools).

.DESCRIPTION
    This script is designed to work together with the companion export script:
      - Export-ADUsers-SignatureData.ps1  -> exports AD user attributes to CSV for signature/profile management
      - Import-ADUsers-SignatureData.ps1  -> imports the same CSV format back and updates AD

    Typical workflow:
      1) Run Export-ADUsers-SignatureData.ps1 to generate a CSV
      2) Edit the CSV in Excel (department, title, phone, address, ExchAttr1-15, etc.)
      3) Run this script to apply ONLY the needed changes back to AD

    Safety rules:
      - Only touches attributes that have a matching column PRESENT in the CSV headers.
      - If a column is NOT in the CSV, the related AD attribute is left unchanged.
      - By default, blank CSV values do NOT clear AD attributes ($AllowClearing = $false).

    CSV robustness:
      - Auto-detects delimiter (comma/semicolon/tab), common when Excel saves CSV with ';' in DK locales.
      - Trims header names to avoid issues with accidental whitespace in Excel.

    Logging:
      - Writes a timestamped .log and a detailed change .csv (planned/applied/failed).

.NOTES
    Author  : Peter
    Script  : Import-ADUsers-SignatureData.ps1
    Version : 1.7
    Updated : 2025-12-15
    Output  : Defaults to .\Logs (for logs)

.REQUIREMENTS
    - RSAT ActiveDirectory module
    - CSV must include at least one identity column: SamAccountName OR UPN OR Email
#>

# -----------------------------
# Make relative paths predictable
# -----------------------------
try { Set-Location -Path $PSScriptRoot } catch { }

# -----------------------------
# Config
# -----------------------------
$CsvPath           = ".\AD_Users_SignatureData_import_us.csv"
$LogDirectory      = ".\Logs"

$WhatIfMode        = $false     # true = log intended changes only (no Set-ADUser calls)
$AllowClearing     = $false     # true = if CSV field is blank (AND column exists), clear the AD attribute

$UpdateMailAttr    = $true      # only if Email column exists
$UpdateProxyAddrs  = $true      # only if Email/ProxyAddresses columns exist
$UpdateManager     = $true      # only if Manager / Manager Email columns exist
$AddCsvProxyAddrs  = $true      # add addresses from CSV "ProxyAddresses" (never removes existing)

# Override delimiter if you want (',' ';' "`t"), else auto-detect
$CsvDelimiter      = $null

# Match order for identifying the user in AD
$IdentityMatchOrder = @('SamAccountName','UPN','Email')

# -----------------------------
# Module + Logging Prep
# -----------------------------
Import-Module ActiveDirectory -ErrorAction Stop

if (-not (Test-Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$RunTs       = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$LogTextPath = Join-Path $LogDirectory "Import-ADUsers-SignatureData_$RunTs.log"
$LogCsvPath  = Join-Path $LogDirectory "Import-ADUsers-SignatureData_Changes_$RunTs.csv"

function Write-Log {
    param(
        [Parameter(Mandatory)] [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','CHANGE')] [string] $Level = 'INFO'
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts][$Level] $Message"
    Write-Host $line
    Add-Content -Path $LogTextPath -Value $line -Encoding UTF8
}

$changeRows = New-Object System.Collections.Generic.List[object]
function Add-ChangeRow {
    param(
        [string]$SamAccountName,
        [string]$UPN,
        [string]$Attribute,
        [string]$OldValue,
        [string]$NewValue,
        [string]$Action,
        [string]$Status,
        [string]$Note
    )
    $changeRows.Add([pscustomobject]@{
        Timestamp      = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        SamAccountName = $SamAccountName
        UPN            = $UPN
        Attribute      = $Attribute
        OldValue       = $OldValue
        NewValue       = $NewValue
        Action         = $Action
        Status         = $Status
        Note           = $Note
    }) | Out-Null
}

function Normalize-String {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Split-SemicolonList {
    param([string]$Value)
    $v = Normalize-String $Value
    if ([string]::IsNullOrWhiteSpace($v)) { return @() }
    return $v.Split(';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Get-DetectedDelimiter {
    param([string]$Path)

    $lines = Get-Content -Path $Path -TotalCount 50 -ErrorAction Stop
    $header = ($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1)
    if (-not $header) { return ',' }

    $comma = ([regex]::Matches($header, ',')).Count
    $semi  = ([regex]::Matches($header, ';')).Count
    $tab   = ([regex]::Matches($header, "`t")).Count

    $max = ($comma, $semi, $tab | Measure-Object -Maximum).Maximum
    if ($max -eq 0) { return ',' }

    if ($semi -eq $max) { return ';' }
    if ($tab  -eq $max) { return "`t" }
    return ','
}

function Ensure-ProxyAddresses {
    param(
        [string[]]$Existing,
        [string]$PrimaryEmail,
        [string[]]$CsvProxyAddresses
    )

    $existingList = @()
    if ($Existing) { $existingList = @($Existing) }

    $primaryEmail = (Normalize-String $PrimaryEmail).ToLowerInvariant()

    # Enforce primary SMTP if PrimaryEmail provided
    if (-not [string]::IsNullOrWhiteSpace($primaryEmail)) {
        $newList = @()
        foreach ($p in $existingList) {
            if ($p -like 'SMTP:*') { $newList += ('smtp:' + $p.Substring(5)) } else { $newList += $p }
        }

        $desiredPrimary = "SMTP:$primaryEmail"

        # Promote smtp:primary to SMTP:primary
        $newList = $newList | ForEach-Object {
            if ($_.ToLowerInvariant() -eq ("smtp:$primaryEmail")) { $desiredPrimary } else { $_ }
        }

        # Ensure desired primary exists
        if (-not ($newList | Where-Object { $_.ToLowerInvariant() -eq $desiredPrimary.ToLowerInvariant() })) {
            $newList += $desiredPrimary
        }

        $existingList = $newList
    }

    # Add proxies from CSV (never remove)
    foreach ($p in $CsvProxyAddresses) {
        $pp = Normalize-String $p
        if ([string]::IsNullOrWhiteSpace($pp)) { continue }
        if ($pp -notmatch '^(?i)smtp:') { $pp = "smtp:$pp" }

        if (-not ($existingList | Where-Object { $_.ToLowerInvariant() -eq $pp.ToLowerInvariant() })) {
            $existingList += $pp
        }
    }

    # De-dupe (case-insensitive)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deduped = foreach ($p in $existingList) { if ($seen.Add($p)) { $p } }
    return ,$deduped
}

# -----------------------------
# Friendly CSV path validation
# -----------------------------
$ResolvedCsvPath = $null
try { $ResolvedCsvPath = (Resolve-Path -Path $CsvPath -ErrorAction Stop).Path } catch { }

if (-not $ResolvedCsvPath -or -not (Test-Path -Path $ResolvedCsvPath -PathType Leaf)) {
    $cwd = (Get-Location).Path

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "IMPORT FAILED: CSV file not found" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Expected CSV:" -ForegroundColor Yellow
    Write-Host ("  {0}" -f $CsvPath) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Working folder:" -ForegroundColor Yellow
    Write-Host ("  {0}" -f $cwd) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Tips:" -ForegroundColor Yellow
    Write-Host "  - If CsvPath is .\file.csv, it must exist in the working folder above." -ForegroundColor Yellow
    Write-Host "  - Consider using a full path (C:\...\file.csv)." -ForegroundColor Yellow
    Write-Host ""

    $candidates = Get-ChildItem -Path $cwd -Filter "*.csv" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 10

    if ($candidates) {
        Write-Host "CSV files in working folder (newest first):" -ForegroundColor Yellow
        foreach ($c in $candidates) {
            Write-Host ("  {0}  ({1})" -f $c.Name, $c.LastWriteTime) -ForegroundColor Yellow
        }
        Write-Host ""
    }

    try { Write-Log "IMPORT FAILED: CSV file not found: $CsvPath (WorkingFolder: $cwd)" "ERROR" } catch { }
    Write-Host "Stopping. Fix CsvPath and run again." -ForegroundColor Yellow
    Write-Host ""
    return
}

$CsvPath = $ResolvedCsvPath

# -----------------------------
# Load CSV (delimiter auto-detect)
# -----------------------------
if (-not $CsvDelimiter) {
    $CsvDelimiter = Get-DetectedDelimiter -Path $CsvPath
}

Write-Log ("Starting import. CSV: {0}" -f $CsvPath)
Write-Log ("Delimiter detected/used: '{0}'" -f ($CsvDelimiter -replace "`t", "\t"))
Write-Log ("WhatIfMode={0} | AllowClearing={1} | UpdateMailAttr={2} | UpdateProxyAddrs={3} | UpdateManager={4}" -f $WhatIfMode, $AllowClearing, $UpdateMailAttr, $UpdateProxyAddrs, $UpdateManager)

$rows = Import-Csv -Path $CsvPath -Delimiter $CsvDelimiter
$rows = @($rows)

if (-not $rows -or $rows.Count -eq 0) {
    Write-Log ("IMPORT FAILED: CSV is empty or could not be parsed: {0}" -f $CsvPath) "ERROR"
    Write-Host ""
    Write-Host "IMPORT FAILED: CSV is empty or could not be parsed." -ForegroundColor Yellow
    Write-Host ("  {0}" -f $CsvPath) -ForegroundColor Yellow
    Write-Host ""
    return
}

$rowCount = $rows.Count
if ($rowCount -lt 1) { $rowCount = 1 }

# Header map (trimmed header -> actual property name)
$HeaderMap = @{}
foreach ($p in $rows[0].PSObject.Properties.Name) {
    $trim = $p.Trim()
    if (-not $HeaderMap.ContainsKey($trim)) {
        $HeaderMap[$trim] = $p
    }
}

function Csv-HasColumn {
    param([string]$Name)
    return $HeaderMap.ContainsKey($Name)
}

function Get-CsvValue {
    param(
        [pscustomobject]$Row,
        [string]$ColumnName
    )
    if (-not (Csv-HasColumn $ColumnName)) { return "" }
    $actual = $HeaderMap[$ColumnName]
    return (Normalize-String $Row.$actual)
}

# Validate identity columns exist
$hasAnyIdColumn = (Csv-HasColumn 'SamAccountName') -or (Csv-HasColumn 'UPN') -or (Csv-HasColumn 'Email')
if (-not $hasAnyIdColumn) {
    $found = ($rows[0].PSObject.Properties.Name | Sort-Object) -join ", "
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "IMPORT FAILED: Missing identity columns" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Expected at least one of these columns:" -ForegroundColor Yellow
    Write-Host "  SamAccountName, UPN, Email" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Columns found:" -ForegroundColor Yellow
    Write-Host ("  {0}" -f $found) -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Most common cause: wrong delimiter (Excel often uses ';' in DK)." -ForegroundColor Yellow
    Write-Host ("Delimiter used: '{0}'" -f ($CsvDelimiter -replace "
