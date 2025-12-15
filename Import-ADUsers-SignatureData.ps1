<#
.SYNOPSIS
    Import AD user profile fields from a CSV (same format as Export-ADUsers-SignatureData) and update only what changed.

.DESCRIPTION
    Safety rule:
      - Only touches attributes that have a matching column PRESENT in the CSV headers.
      - If a column is NOT in the CSV, that attribute is left unchanged.
    By default:
      - Does NOT clear attributes when CSV value is blank ($AllowClearing = $false).
    Logging:
      - Writes a timestamped .log and a detailed change .csv (planned/applied/failed).

.NOTES
    Author  : Peter
    Script  : Import-ADUsers-SignatureData.ps1
    Version : 1.4
    Updated : 2025-12-15
    Output  : Defaults to .\Logs (for logs)

.REQUIREMENTS
    - RSAT ActiveDirectory module
    - CSV must include at least: SamAccountName OR UPN OR Email
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
$UpdateManager     = $true      # only if Manager or Manager Email columns exist
$AddCsvProxyAddrs  = $true      # add addresses from CSV "ProxyAddresses" (never removes existing)

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

function Resolve-TargetUser {
    param([pscustomobject]$Row)

    foreach ($key in $IdentityMatchOrder) {
        $val = Normalize-String $Row.$key
        if ([string]::IsNullOrWhiteSpace($val)) { continue }

        try {
            switch ($key) {
                'SamAccountName' {
                    return (Get-ADUser -Identity $val -Properties * -ErrorAction Stop)
                }
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
    param([pscustomobject]$Row, [hashtable]$HeaderSet)

    if ($HeaderSet.ContainsKey('Manager Email')) {
        $mgrEmail = Normalize-String $Row.'Manager Email'
        if (-not [string]::IsNullOrWhiteSpace($mgrEmail)) {
            $m = Get-ADUser -Filter "mail -eq '$mgrEmail'" -Properties DistinguishedName -ErrorAction SilentlyContinue
            if ($m) { return $m.DistinguishedName }
            return $null
        }
    }

    if ($HeaderSet.ContainsKey('Manager')) {
        $mgrName = Normalize-String $Row.'Manager'
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

    # Ensure only ONE primary SMTP, and it matches Email (if Email column exists/provided)
    if (-not [string]::IsNullOrWhiteSpace($primaryEmail)) {
        $newList = @()
        foreach ($p in $existingList) {
            if ($p -like 'SMTP:*') { $newList += ('smtp:' + ($p.Substring(5))) } else { $newList += $p }
        }

        $desiredPrimary = "SMTP:$primaryEmail"

        # If desired exists as smtp:, promote it
        $newList = $newList | ForEach-Object {
            if ($_.ToLowerInvariant() -eq ("smtp:$primaryEmail")) { $desiredPrimary } else { $_ }
        }

        # Ensure desired primary exists
        if (-not ($newList | Where-Object { $_.ToLowerInvariant() -eq $desiredPrimary.ToLowerInvariant() })) {
            $newList += $desiredPrimary
        }

        $existingList = $newList
    }

    # Add any CSV proxies (never remove)
    foreach ($p in $CsvProxyAddresses) {
        $pp = Normalize-String $p
        if ([string]::IsNullOrWhiteSpace($pp)) { continue }
        if ($pp -notmatch '^(?i)smtp:') { $pp = "smtp:$pp" }

        if (-not ($existingList | Where-Object { $_.ToLowerInvariant() -eq $pp.ToLowerInvariant() })) {
            $existingList += $pp
        }
    }

    # De-dupe case-insensitive
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deduped = foreach ($p in $existingList) { if ($seen.Add($p)) { $p } }

    return ,$deduped
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
    Write-Host "  - If you use a relative path like .\file.csv, it is relative to the working folder above."
    Write-Host "  - Try setting CsvPath to a full path, e.g. C:\Scripts\ExportADinfo\file.csv"
    Write-Host ""

    $candidates = Get-ChildItem -Path $cwd -Filter "*.csv" -File -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending |
                  Select-Object -First 10

    if ($candidates) {
        Write-Host "CSV files found in the working folder (newest first):" -ForegroundColor Yellow
        foreach ($c in $candidates) {
            Write-Host ("  {0}  ({1})" -f $c.Name, $c.LastWriteTime) -ForegroundColor Yellow
        }
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
# Load CSV
# -----------------------------
Write-Log "Starting import. CSV: $CsvPath"
Write-Log "WhatIfMode=$WhatIfMode | AllowClearing=$AllowClearing | UpdateMailAttr=$UpdateMailAttr | UpdateProxyAddrs=$UpdateProxyAddrs | UpdateManager=$UpdateManager"

$rows = Import-Csv -Path $CsvPath
$rows = @($rows)  # ensure array even for single-row CSV
if (-not $rows -or $rows.Count -eq 0) {
    Write-Log "IMPORT FAILED: CSV is empty: $CsvPath" "ERROR"
    Write-Host ""
    Write-Host "IMPORT FAILED: The CSV file is empty:" -ForegroundColor Yellow
    Write-Host "  $CsvPath" -ForegroundColor Yellow
    Write-Host ""
    return
}

$rowCount = $rows.Count
if ($rowCount -lt 1) { $rowCount = 1 }

# Build CSV header set
$HeaderSet = @{}
$rows[0].PSObject.Properties.Name | ForEach-Object { $HeaderSet[$_] = $true }

function Csv-HasColumn {
    param([string]$Name)
    return $HeaderSet.ContainsKey($Name)
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

    $idDisplay = Normalize-String $r.SamAccountName
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Normalize-String $r.UPN }
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Normalize-String $r.Email }
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
            Add-ChangeRow -SamAccountName (Normalize-String $r.SamAccountName) -UPN (Normalize-String $r.UPN) -Attribute "(identity)" -OldValue "" -NewValue "" -Action "Skip" -Status "NotFound" -Note "No matching AD user"
            continue
        }

        $matched++

        $replace = @{}
        $clear   = @()
        $planSetManager = $false
        $mgrDn = $null

        # Map CSV columns -> AD attribute names (ONLY if column exists)
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

            $csvVal = Normalize-String $r.($m.Csv)
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

        # otherTelephone (multi-valued) only if column exists
        if (Csv-HasColumn 'Other phones') {
            $csvOtherPhones = Split-SemicolonList (Normalize-String $r.'Other phones')
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

        # Email / mail only if Email column exists
        $csvEmail = ""
        if (Csv-HasColumn 'Email') { $csvEmail = Normalize-String $r.Email }

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

        # proxyAddresses only if Email/ProxyAddresses columns exist
        $proxyRelevant = (Csv-HasColumn 'Email') -or (Csv-HasColumn 'ProxyAddresses')
        if ($UpdateProxyAddrs -and $proxyRelevant) {
            $existingProxies = @()
            if ($u.proxyAddresses) { $existingProxies = @($u.proxyAddresses) }

            $csvProxies = @()
            if ($AddCsvProxyAddrs -and (Csv-HasColumn 'ProxyAddresses')) {
                $csvProxies = Split-SemicolonList (Normalize-String $r.ProxyAddresses)
            }

            $newProxies = Ensure-ProxyAddresses -Existing $existingProxies -PrimaryEmail $csvEmail -CsvProxyAddresses $csvProxies

            $oldNorm = ($existingProxies | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'
            $newNorm = ($newProxies      | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'

            if ($oldNorm -ne $newNorm) {
                $replace['proxyAddresses'] = $newProxies
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'proxyAddresses' -OldValue ($existingProxies -join ';') -NewValue ($newProxies -join ';') -Action "Replace" -Status "Planned" -Note "Keeps existing; enforces primary SMTP only if Email column exists"
            }
        }

        # manager only if relevant columns exist
        $mgrRelevant = (Csv-HasColumn 'Manager') -or (Csv-HasColumn 'Manager Email')
        if ($UpdateManager -and $mgrRelevant) {
            $mgrDn = Resolve-ManagerDn -Row $r -HeaderSet $HeaderSet
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

        # Apply Replace/Clear
        if ($replace.Count -gt 0 -or $clear.Count -gt 0) {
            $params = @{ Identity = $u.DistinguishedName; ErrorAction = 'Stop' }
            if ($replace.Count -gt 0) { $params['Replace'] = $replace }
            if ($clear.Count   -gt 0) { $params['Clear']   = $clear }
            Set-ADUser @params
        }

        # Apply Manager
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
        Add-ChangeRow -SamAccountName (Normalize-String $r.SamAccountName) -UPN (Normalize-String $r.UPN) -Attribute "(row)" -OldValue "" -NewValue "" -Action "Error" -Status "Failed" -Note $($_.Exception.Message)
    }
}

Write-Progress -Activity "Importing users" -Completed

# -----------------------------
# Write change CSV
# -----------------------------
if ($changeRows.Count -gt 0) {
    $changeRows | Export-Csv -Path $LogCsvPath -NoTypeInformation -Encoding UTF8
} else {
    "Timestamp,SamAccountName,UPN,Attribute,OldValue,NewValue,Action,Status,Note" | Set-Content -Path $LogCsvPath -Encoding UTF8
}

# -----------------------------
# Summary
# -----------------------------
Write-Log "Run complete."
Write-Log "Processed : $processed"
Write-Log "Matched   : $matched"
Write-Log "Updated   : $updated"
Write-Log "Skipped   : $skipped"
Write-Log "Errors    : $errors"
Write-Log "Log (text): $LogTextPath"
Write-Log "Log (csv) : $LogCsvPath"
