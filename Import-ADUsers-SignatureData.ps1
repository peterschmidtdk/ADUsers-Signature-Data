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

    Safety rule:
      - Only touches attributes that have a matching column PRESENT in the CSV headers.
      - If a column is NOT in the CSV, that attribute is left unchanged.
      - By default, blank CSV values do NOT clear AD attributes ($AllowClearing = $false).

    CSV robustness:
      - Auto-detects delimiter (comma/semicolon/tab) to handle Excel regional CSV formats.
      - Trims header names to avoid issues with accidental whitespace in column names.

    Logging:
      - Writes a timestamped .log and a detailed change .csv (planned/applied/failed) for traceability.

.NOTES
    Author  : Peter
    Script  : Import-ADUsers-SignatureData.ps1
    Version : 1.6
    Updated : 2025-12-15
    Output  : Defaults to .\Logs (for logs)

.REQUIREMENTS
    - RSAT ActiveDirectory module
    - CSV must include at least: SamAccountName OR UPN OR Email
#>

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
$UpdateManager     = $true      # only if Manager or Manager Email columns exist
$AddCsvProxyAddrs  = $true      # add addresses from CSV "ProxyAddresses" (never removes existing)

# If you want to override autodetect, set explicitly e.g. ';' or ','
$CsvDelimiter      = $null

# Match order for identifying the user in AD
$IdentityMatchOrder = @('SamAccountName','UPN','Email')

# -----------------------------
# Prep
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

    # Read the first non-empty line (header)
    $lines = Get-Content -Path $Path -TotalCount 50 -ErrorAction Stop
    $header = ($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1)

    if (-not $header) { return ',' }

    $comma = ([regex]::Matches($header, ',')).Count
    $semi  = ([regex]::Matches($header, ';')).Count
    $tab   = ([regex]::Matches($header, "`t")).Count

    # Pick the delimiter with the most occurrences
    $max = ($comma, $semi, $tab | Measure-Object -Maximum).Maximum
    if ($max -eq 0) { return ',' }

    if ($semi -eq $max) { return ';' }
    if ($tab  -eq $max) { return "`t" }
    return ','
}

# -----------------------------
# Friendly CSV path validation (no stack trace)
# -----------------------------
$ResolvedCsvPath = $null
try { $ResolvedCsvPath = (Resolve-Path -Path $CsvPath -ErrorAction Stop).Path } catch { }

if (-not $ResolvedCsvPath -or -not (Test-Path -Path $ResolvedCsvPath -PathType Leaf)) {

    $cwd = (Get-Location).Path

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "IMPORT FAILED: CSV file not found" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "The script expected this CSV file:" -ForegroundColor Yellow
    Write-Host "  $CsvPath" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Current working folder:" -ForegroundColor Yellow
    Write-Host "  $cwd" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Tips:" -ForegroundColor Yellow
    Write-Host "  - If you use .\file.csv, it is relative to the working folder above."
    Write-Host "  - Try setting CsvPath to a full path, e.g. C:\Scripts\ExportADinfo\file.csv"
    Write-Host ""

    $candidates = Get-ChildItem -Path $cwd -Filter "*.csv" -File -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending |
                  Select-Object -First 10

    if ($candidates) {
        Write-Host "CSV files found in the working folder (newest first):" -ForegroundColor Yellow
        foreach ($c in $candidates) { Write-Host ("  {0}  ({1})" -f $c.Name, $c.LastWriteTime) -ForegroundColor Yellow }
        Write-Host ""
    } else {
        Write-Host "No CSV files were found in the working folder." -ForegroundColor Yellow
        Write-Host ""
    }

    try { Write-Log "IMPORT FAILED: CSV file not found: $CsvPath (WorkingFolder: $cwd)" "ERROR" } catch { }
    Write-Host "Stopping. Fix CsvPath and run again." -ForegroundColor Yellow
    Write-Host ""
    return
}

$CsvPath = $ResolvedCsvPath

# -----------------------------
# Load CSV (with delimiter autodetect)
# -----------------------------
if (-not $CsvDelimiter) {
    $CsvDelimiter = Get-DetectedDelimiter -Path $CsvPath
}

Write-Log "Starting import. CSV: $CsvPath"
Write-Log ("Delimiter detected/used: '{0}'" -f ($CsvDelimiter -replace "`t", "\t"))
Write-Log "WhatIfMode=$WhatIfMode | AllowClearing=$AllowClearing | UpdateMailAttr=$UpdateMailAttr | UpdateProxyAddrs=$UpdateProxyAddrs | UpdateManager=$UpdateManager"

$rows = Import-Csv -Path $CsvPath -Delimiter $CsvDelimiter
$rows = @($rows)

if (-not $rows -or $rows.Count -eq 0) {
    Write-Log "IMPORT FAILED: CSV is empty or could not be parsed: $CsvPath" "ERROR"
    Write-Host ""
    Write-Host "IMPORT FAILED: The CSV file is empty (or could not be parsed):" -ForegroundColor Yellow
    Write-Host "  $CsvPath" -ForegroundColor Yellow
    Write-Host ""
    return
}

$rowCount = $rows.Count
if ($rowCount -lt 1) { $rowCount = 1 }

# Build header map (trimmed header -> actual property name)
$HeaderMap = @{}
foreach ($p in $rows[0].PSObject.Properties.Name) {
    $trim = ($p.Trim())
    if (-not $HeaderMap.ContainsKey($trim)) { $HeaderMap[$trim] = $p }
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
    return Normalize-String $Row.$actual
}

# Validate identity columns exist in header
$hasAnyIdColumn = (Csv-HasColumn 'SamAccountName') -or (Csv-HasColumn 'UPN') -or (Csv-HasColumn 'Email')
if (-not $hasAnyIdColumn) {
    $found = ($rows[0].PSObject.Properties.Name | Sort-Object) -join ", "
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "IMPORT FAILED: Missing identity columns" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host "Expected at least one of these columns in the CSV header:" -ForegroundColor Yellow
    Write-Host "  SamAccountName, UPN, Email" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Columns found in your file:" -ForegroundColor Yellow
    Write-Host "  $found" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Most common cause: wrong CSV delimiter (Excel often saves with ';' in DK)." -ForegroundColor Yellow
    Write-Host ("Delimiter used was: '{0}'" -f ($CsvDelimiter -replace "`t", "\t")) -ForegroundColor Yellow
    Write-Host ""
    Write-Log "IMPORT FAILED: No identity columns found. ColumnsFound=[$found]" "ERROR"
    return
}

function Resolve-TargetUser {
    param([pscustomobject]$Row)

    foreach ($key in $IdentityMatchOrder) {
        $val = Get-CsvValue -Row $Row -ColumnName $key
        if ([string]::IsNullOrWhiteSpace($val)) { continue }

        try {
            switch ($key) {
                'SamAccountName' { return (Get-ADUser -Identity $val -Properties * -ErrorAction Stop) }
                'UPN' {
                    $u = Get-ADUser -Filter "UserPrincipalName -eq '$val'" -Properties * -ErrorAction Stop
                    if ($u) { return $u }
                }
                'Email' {
                    $u = Get-ADUser -Filter "mail -eq '$val'" -Properties * -ErrorAction Stop
                    if ($u) { return $u }
                }
            }
        } catch { }
    }
    return $null
}

function Resolve-ManagerDn {
    param([pscustomobject]$Row)

    if (Csv-HasColumn 'Manager Email') {
        $mgrEmail = Get-CsvValue -Row $Row -ColumnName 'Manager Email'
        if (-not [string]::IsNullOrWhiteSpace($mgrEmail)) {
            $m = Get-ADUser -Filter "mail -eq '$mgrEmail'" -Properties DistinguishedName -ErrorAction SilentlyContinue
            if ($m) { return $m.DistinguishedName }
            return $null
        }
    }

    if (Csv-HasColumn 'Manager') {
        $mgrName = Get-CsvValue -Row $Row -ColumnName 'Manager'
        if (-not [string]::IsNullOrWhiteSpace($mgrName)) {
            $cands = Get-ADUser -Filter "DisplayName -eq '$mgrName'" -Properties DistinguishedName -ErrorAction SilentlyContinue
            if ($cands.Count -eq 1) { return $cands[0].DistinguishedName }
            return $null
        }
    }

    return $null
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

    if (-not [string]::IsNullOrWhiteSpace($primaryEmail)) {
        $newList = @()
        foreach ($p in $existingList) {
            if ($p -like 'SMTP:*') { $newList += ('smtp:' + ($p.Substring(5))) } else { $newList += $p }
        }

        $desiredPrimary = "SMTP:$primaryEmail"
        $newList = $newList | ForEach-Object {
            if ($_.ToLowerInvariant() -eq ("smtp:$primaryEmail")) { $desiredPrimary } else { $_ }
        }

        if (-not ($newList | Where-Object { $_.ToLowerInvariant() -eq $desiredPrimary.ToLowerInvariant() })) {
            $newList += $desiredPrimary
        }

        $existingList = $newList
    }

    foreach ($p in $CsvProxyAddresses) {
        $pp = Normalize-String $p
        if ([string]::IsNullOrWhiteSpace($pp)) { continue }
        if ($pp -notmatch '^(?i)smtp:') { $pp = "smtp:$pp" }

        if (-not ($existingList | Where-Object { $_.ToLowerInvariant() -eq $pp.ToLowerInvariant() })) {
            $existingList += $pp
        }
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deduped = foreach ($p in $existingList) { if ($seen.Add($p)) { $p } }

    return ,$deduped
}

# -----------------------------
# Import loop
# -----------------------------
$processed = 0
$matched   = 0
$updated   = 0
$skipped   = 0
$errors    = 0

foreach ($r in $rows) {
    $processed++

    $idDisplay = Get-CsvValue -Row $r -ColumnName 'SamAccountName'
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Get-CsvValue -Row $r -ColumnName 'UPN' }
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Get-CsvValue -Row $r -ColumnName 'Email' }
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = "(no id in row)" }

    $pct = [int](($processed / $rowCount) * 100)
    if ($pct -gt 100) { $pct = 100 }
    if ($pct -lt 0)   { $pct = 0 }

    Write-Progress -Activity "Importing users" -Status ("Processing {0} / {1}: {2}" -f $processed, $rowCount, $idDisplay) -PercentComplete $pct
    Write-Host ("[{0}/{1}] {2}" -f $processed, $rowCount, $idDisplay)

    try {
        $u = Resolve-TargetUser -Row $r
        if (-not $u) {
            $skipped++
            Write-Log "User not found in AD for row identity: $idDisplay" "WARN"
            Add-ChangeRow -SamAccountName (Get-CsvValue -Row $r -ColumnName 'SamAccountName') -UPN (Get-CsvValue -Row $r -ColumnName 'UPN') -Attribute "(identity)" -OldValue "" -NewValue "" -Action "Skip" -Status "NotFound" -Note "No matching AD user"
            continue
        }

        $matched++

        $replace = @{}
        $clear   = @()
        $planSetManager = $false
        $mgrDn = $null

        $map = @(
            @{ Csv='First Name';     Ad='givenName' },
            @{ Csv='Last Name';      Ad='sn' },
            @{ Csv='Initials';       Ad='initials' },

            @{ Csv='Company';        Ad='company' },
            @{ Csv='Department';     Ad='department' },
            @{ Csv='Office';         Ad='physicalDeliveryOfficeName' },
            @{ Csv='Position';       Ad='title' },
            @{ Csv='Description';    Ad='description' },
            @{ Csv='Notes';          Ad='info' },

            @{ Csv='Street';         Ad='streetAddress' },
            @{ Csv='P.O. Box';       Ad='postOfficeBox' },
            @{ Csv='City';           Ad='l' },
            @{ Csv='State';          Ad='st' },
            @{ Csv='Postal code';    Ad='postalCode' },
            @{ Csv='Country (c)';    Ad='c' },
            @{ Csv='Country (co)';   Ad='co' },

            @{ Csv='Phone';          Ad='telephoneNumber' },
            @{ Csv='Mobile';         Ad='mobile' },
            @{ Csv='IP phone';       Ad='ipPhone' },
            @{ Csv='Home phone';     Ad='homePhone' },
            @{ Csv='Fax';            Ad='facsimileTelephoneNumber' },
            @{ Csv='Pager';          Ad='pager' },
            @{ Csv='Web page';       Ad='wWWHomePage' },

            @{ Csv='EmployeeID';     Ad='employeeID' },
            @{ Csv='EmployeeNumber'; Ad='employeeNumber' },
            @{ Csv='EmployeeType';   Ad='employeeType' },

            @{ Csv='ExchAttr1';      Ad='extensionAttribute1' },
            @{ Csv='ExchAttr2';      Ad='extensionAttribute2' },
            @{ Csv='ExchAttr3';      Ad='extensionAttribute3' },
            @{ Csv='ExchAttr4';      Ad='extensionAttribute4' },
            @{ Csv='ExchAttr5';      Ad='extensionAttribute5' },
            @{ Csv='ExchAttr6';      Ad='extensionAttribute6' },
            @{ Csv='ExchAttr7';      Ad='extensionAttribute7' },
            @{ Csv='ExchAttr8';      Ad='extensionAttribute8' },
            @{ Csv='ExchAttr9';      Ad='extensionAttribute9' },
            @{ Csv='ExchAttr10';     Ad='extensionAttribute10' },
            @{ Csv='ExchAttr11';     Ad='extensionAttribute11' },
            @{ Csv='ExchAttr12';     Ad='extensionAttribute12' },
            @{ Csv='ExchAttr13';     Ad='extensionAttribute13' },
            @{ Csv='ExchAttr14';     Ad='extensionAttribute14' },
            @{ Csv='ExchAttr15';     Ad='extensionAttribute15' }
        )

        foreach ($m in $map) {
            if (-not (Csv-HasColumn $m.Csv)) { continue }

            $csvVal = Get-CsvValue -Row $r -ColumnName $m.Csv
            $adVal  = Normalize-String $u.($m.Ad)

            if ([string]::IsNullOrWhiteSpace($csvVal)) {
                if ($AllowClearing -and -not [string]::IsNullOrWhiteSpace($adVal)) {
                    $clear += $m.Ad
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute $m.Ad -OldValue $adVal -NewValue "" -Action "Clear" -Status "Planned" -Note "CSV blank and AllowClearing enabled"
                }
                continue
            }

            if ($csvVal -ne $adVal) {
                $replace[$m.Ad] = $csvVal
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute $m.Ad -OldValue $adVal -NewValue $csvVal -Action "Replace" -Status "Planned" -Note ""
            }
        }

        # otherTelephone
        if (Csv-HasColumn 'Other phones') {
            $csvOtherPhones = Split-SemicolonList (Get-CsvValue -Row $r -ColumnName 'Other phones')
            $adOtherPhones  = @()
            if ($u.otherTelephone) { $adOtherPhones = @($u.otherTelephone) }

            if ($csvOtherPhones.Count -gt 0) {
                $csvSet = $csvOtherPhones | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object
                $adSet  = $adOtherPhones  | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object

                if (($csvSet -join '|') -ne ($adSet -join '|')) {
                    $replace['otherTelephone'] = $csvOtherPhones
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'otherTelephone' -OldValue ($adOtherPhones -join ';') -NewValue ($csvOtherPhones -join ';') -Action "Replace" -Status "Planned" -Note "Multi-valued"
                }
            } elseif ($AllowClearing -and $adOtherPhones.Count -gt 0) {
                $clear += 'otherTelephone'
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'otherTelephone' -OldValue ($adOtherPhones -join ';') -NewValue "" -Action "Clear" -Status "Planned" -Note "CSV blank and AllowClearing enabled"
            }
        }

        # mail
        $csvEmail = ""
        if (Csv-HasColumn 'Email') { $csvEmail = Get-CsvValue -Row $r -ColumnName 'Email' }

        if ($UpdateMailAttr -and (Csv-HasColumn 'Email')) {
            if (-not [string]::IsNullOrWhiteSpace($csvEmail)) {
                $adMail = Normalize-String $u.mail
                if ($csvEmail -ne $adMail) {
                    $replace['mail'] = $csvEmail
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'mail' -OldValue $adMail -NewValue $csvEmail -Action "Replace" -Status "Planned" -Note ""
                }
            } elseif ($AllowClearing -and -not [string]::IsNullOrWhiteSpace((Normalize-String $u.mail))) {
                $clear += 'mail'
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'mail' -OldValue (Normalize-String $u.mail) -NewValue "" -Action "Clear" -Status "Planned" -Note "CSV blank and AllowClearing enabled"
            }
        }

        # proxyAddresses
        $proxyRelevant = (Csv-HasColumn 'Email') -or (Csv-HasColumn 'ProxyAddresses')
        if ($UpdateProxyAddrs -and $proxyRelevant) {
            $existingProxies = @()
            if ($u.proxyAddresses) { $existingProxies = @($u.proxyAddresses) }

            $csvProxies = @()
            if ($AddCsvProxyAddrs -and (Csv-HasColumn 'ProxyAddresses')) {
                $csvProxies = Split-SemicolonList (Get-CsvValue -Row $r -ColumnName 'ProxyAddresses')
            }

            $newProxies = Ensure-ProxyAddresses -Existing $existingProxies -PrimaryEmail $csvEmail -CsvProxyAddresses $csvProxies

            $oldNorm = ($existingProxies | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'
            $newNorm = ($newProxies      | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'

            if ($oldNorm -ne $newNorm) {
                $replace['proxyAddresses'] = $newProxies
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'proxyAddresses' -OldValue ($existingProxies -join ';') -NewValue ($newProxies -join ';') -Action "Replace" -Status "Planned" -Note "Keeps existing; enforces primary SMTP only if Email column exists"
            }
        }

        # manager
        $mgrRelevant = (Csv-HasColumn 'Manager') -or (Csv-HasColumn 'Manager Email')
        if ($UpdateManager -and $mgrRelevant) {
            $mgrDn = Resolve-ManagerDn -Row $r
            if ($mgrDn) {
                $oldMgr = Normalize-String $u.manager
                if ($mgrDn -ne $oldMgr) {
                    $planSetManager = $true
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'manager' -OldValue $oldMgr -NewValue $mgrDn -Action "Set" -Status "Planned" -Note "Resolved manager DN"
                }
            } else {
                Write-Log "Manager could not be uniquely resolved for $($u.SamAccountName). Skipping manager update." "WARN"
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'manager' -OldValue (Normalize-String $u.manager) -NewValue "" -Action "Skip" -Status "Unresolved" -Note "Manager not uniquely resolvable"
            }
        }

        $hasChanges = ($replace.Count -gt 0) -or ($clear.Count -gt 0) -or $planSetManager

        if (-not $hasChanges) {
            $skipped++
            Write-Log "No changes needed for $($u.SamAccountName)."
            continue
        }

        if ($WhatIfMode) {
            $updated++
            Write-Log "WhatIfMode: would update $($u.SamAccountName) with $($replace.Count) replace(s), $($clear.Count) clear(s), manager=$planSetManager." "CHANGE"
            continue
        }

        if ($replace.Count -gt 0 -or $clear.Count -gt 0) {
            $params = @{ Identity = $u.DistinguishedName; ErrorAction = 'Stop' }
            if ($replace.Count -gt 0) { $params['Replace'] = $replace }
            if ($clear.Count   -gt 0) { $params['Clear']   = $clear }
            Set-ADUser @params
        }

        if ($planSetManager -and $mgrDn) {
            Set-ADUser -Identity $u.DistinguishedName -Manager $mgrDn -ErrorAction Stop
        }

        $updated++
        Write-Log "Updated $($u.SamAccountName)." "CHANGE"

        foreach ($cr in $changeRows | Where-Object { $_.SamAccountName -eq $u.SamAccountName -and $_.Status -eq 'Planned' }) {
            $cr.Status = 'Applied'
        }
    }
    catch {
        $errors++
        Write-Log "Error processing $idDisplay : $($_.Exception.Message)" "ERROR"
        Add-ChangeRow -SamAccountName (Get-CsvValue -Row $r -ColumnName 'SamAccountName') -UPN (Get-CsvValue -Row $r -ColumnName 'UPN') -Attribute "(row)" -OldValue "" -NewValue "" -Acti
