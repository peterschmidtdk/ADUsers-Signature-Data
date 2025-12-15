<#
.SYNOPSIS
    Import AD user profile fields from a CSV (same format as Export-ADUsers-SignatureData) and update only what changed.

.DESCRIPTION
    Reads a CSV with the same headers produced by the export script and updates corresponding AD attributes.
    - Updates ONLY attributes that differ (change detection per attribute).
    - By default: does NOT clear attributes when CSV value is blank (safer). Toggle $AllowClearing to change this.
    - ProxyAddresses: keeps existing addresses, ensures a single Primary SMTP (SMTP:) matches CSV "Email",
      and optionally adds any addresses from CSV "ProxyAddresses" (does not remove extras).
    - Manager: resolves from "Manager Email" (preferred) or "Manager" (DisplayName) if unique.

.NOTES
    Author  : Peter
    Script  : Import-ADUsers-SignatureData.ps1
    Version : 1.0
    Updated : 2025-12-15
    Output  : Defaults to .\

.REQUIREMENTS
    - RSAT ActiveDirectory module
    - CSV must include at least: SamAccountName OR UPN OR Email
#>

# -----------------------------
# Config
# -----------------------------
$CsvPath           = ".\AD_Users_SignatureData_Export_YYYY-MM-dd_HH-mm-ss.csv"
$LogDirectory      = ".\Logs"
$WhatIfMode        = $false     # true = log intended changes only (no Set-ADUser calls)
$AllowClearing     = $false     # true = if CSV field is blank, clear the AD attribute
$UpdateMailAttr    = $true      # true = set AD 'mail' to CSV Email (when different)
$UpdateProxyAddrs  = $true      # true = update proxyAddresses logic
$UpdateManager     = $true      # true = try to set manager from CSV
$AddCsvProxyAddrs  = $true      # true = add addresses from CSV "ProxyAddresses" (never removes existing)

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

# Change log rows (CSV)
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
                    $u = Get-ADUser -Identity $val -Properties * -ErrorAction Stop
                    return $u
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
        } catch {
            # keep trying next match type
        }
    }
    return $null
}

function Resolve-ManagerDn {
    param([pscustomobject]$Row)

    $mgrEmail = Normalize-String $Row.'Manager Email'
    if (-not [string]::IsNullOrWhiteSpace($mgrEmail)) {
        $m = Get-ADUser -Filter "mail -eq '$mgrEmail'" -Properties DistinguishedName -ErrorAction SilentlyContinue
        if ($m) { return $m.DistinguishedName }
        return $null
    }

    $mgrName = Normalize-String $Row.'Manager'
    if (-not [string]::IsNullOrWhiteSpace($mgrName)) {
        $cands = Get-ADUser -Filter "DisplayName -eq '$mgrName'" -Properties DistinguishedName -ErrorAction SilentlyContinue
        if ($cands.Count -eq 1) { return $cands[0].DistinguishedName }
        return $null
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

    # Helper: normalize "smtp:" casing for comparisons
    function NormProxy([string]$p) { return ($p.Trim()) }

    $primaryEmail = (Normalize-String $PrimaryEmail).ToLowerInvariant()

    # 1) Ensure only ONE primary SMTP and it matches PrimaryEmail (if provided)
    if (-not [string]::IsNullOrWhiteSpace($primaryEmail)) {
        $newList = @()

        # Demote any existing primary SMTP
        foreach ($p in $existingList) {
            if ($p -like 'SMTP:*') {
                $newList += ('smtp:' + ($p.Substring(5)))
            } else {
                $newList += $p
            }
        }

        # Promote/ensure desired primary
        $desiredPrimary = "SMTP:$primaryEmail"
        $already = $newList | Where-Object { $_.ToLowerInvariant() -eq $desiredPrimary.ToLowerInvariant() }
        if ($already) {
            # If exists as smtp:, replace with SMTP:
            $newList = $newList | ForEach-Object {
                if ($_.ToLowerInvariant() -eq ("smtp:$primaryEmail")) { $desiredPrimary } else { $_ }
            }
        } else {
            $newList += $desiredPrimary
        }

        $existingList = $newList
    }

    # 2) Optionally add any proxy addresses present in CSV (never removes)
    foreach ($p in $CsvProxyAddresses) {
        $pp = Normalize-String $p
        if ([string]::IsNullOrWhiteSpace($pp)) { continue }

        # If it's a bare email, treat as smtp:
        if ($pp -notmatch '^(?i)smtp:') {
            $pp = "smtp:$pp"
        }

        $exists = $existingList | Where-Object { $_.ToLowerInvariant() -eq $pp.ToLowerInvariant() }
        if (-not $exists) { $existingList += $pp }
    }

    # de-dupe (case-insensitive), preserve first occurrence
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deduped = foreach ($p in $existingList) {
        if ($seen.Add($p)) { $p }
    }

    return ,$deduped
}

# -----------------------------
# Validate input
# -----------------------------
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    throw
}

Write-Log "Starting import. CSV: $CsvPath"
Write-Log "WhatIfMode=$WhatIfMode | AllowClearing=$AllowClearing | UpdateMailAttr=$UpdateMailAttr | UpdateProxyAddrs=$UpdateProxyAddrs | UpdateManager=$UpdateManager"

$rows = Import-Csv -Path $CsvPath
if (-not $rows -or $rows.Count -eq 0) {
    Write-Error "CSV is empty: $CsvPath"
    throw
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

    $idDisplay = (Normalize-String $r.SamAccountName)
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Normalize-String $r.UPN }
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = Normalize-String $r.Email }
    if ([string]::IsNullOrWhiteSpace($idDisplay)) { $idDisplay = "(no id in row)" }

    Write-Progress -Activity "Importing users" -Status "Processing $processed / $($rows.Count): $idDisplay" -PercentComplete ([int](($processed / $rows.Count) * 100))
    Write-Log "Processing: $idDisplay"

    try {
        $u = Resolve-TargetUser -Row $r
        if (-not $u) {
            $skipped++
            Write-Log "User not found in AD for row identity: $idDisplay" "WARN"
            Add-ChangeRow -SamAccountName (Normalize-String $r.SamAccountName) -UPN (Normalize-String $r.UPN) -Attribute "(identity)" -OldValue "" -NewValue "" -Action "Skip" -Status "NotFound" -Note "No matching AD user"
            continue
        }

        $matched++

        # Build changes
        $replace = @{}
        $clear   = @()
        $notes   = @()

        # Map CSV columns -> AD attribute names
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

        # Multi-valued: otherTelephone (CSV: "Other phones" ; separated)
        $csvOtherPhones = Split-SemicolonList (Normalize-String $r.'Other phones')
        $adOtherPhones  = @()
        if ($u.otherTelephone) { $adOtherPhones = @($u.otherTelephone) }

        if ($csvOtherPhones.Count -gt 0) {
            # Compare sets case-insensitive
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

        # Email / mail + proxyAddresses
        $csvEmail = Normalize-String $r.Email

        if ($UpdateMailAttr -and -not [string]::IsNullOrWhiteSpace($csvEmail)) {
            $adMail = Normalize-String $u.mail
            if ($csvEmail -ne $adMail) {
                $replace['mail'] = $csvEmail
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'mail' -OldValue $adMail -NewValue $csvEmail -Action "Replace" -Status "Planned" -Note ""
            }
        } elseif ($AllowClearing -and $UpdateMailAttr -and [string]::IsNullOrWhiteSpace($csvEmail) -and -not [string]::IsNullOrWhiteSpace((Normalize-String $u.mail))) {
            $clear += 'mail'
            Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'mail' -OldValue (Normalize-String $u.mail) -NewValue "" -Action "Clear" -Status "Planned" -Note "CSV blank and AllowClearing enabled"
        }

        if ($UpdateProxyAddrs) {
            $existingProxies = @()
            if ($u.proxyAddresses) { $existingProxies = @($u.proxyAddresses) }

            $csvProxies = @()
            if ($AddCsvProxyAddrs) {
                $csvProxies = Split-SemicolonList (Normalize-String $r.ProxyAddresses)
            }

            $newProxies = Ensure-ProxyAddresses -Existing $existingProxies -PrimaryEmail $csvEmail -CsvProxyAddresses $csvProxies

            # Compare sets case-insensitive
            $oldNorm = ($existingProxies | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'
            $newNorm = ($newProxies      | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join '|'

            if ($oldNorm -ne $newNorm) {
                $replace['proxyAddresses'] = $newProxies
                Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'proxyAddresses' -OldValue ($existingProxies -join ';') -NewValue ($newProxies -join ';') -Action "Replace" -Status "Planned" -Note "Keeps existing; enforces single primary SMTP if Email provided"
            }
        }

        # Manager
        if ($UpdateManager) {
            $mgrDn = Resolve-ManagerDn -Row $r
            if ($mgrDn) {
                $oldMgr = Normalize-String $u.manager
                if ($mgrDn -ne $oldMgr) {
                    # Set-ADUser uses -Manager parameter (not Replace) => handle later
                    $notes += "SetManager"
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'manager' -OldValue $oldMgr -NewValue $mgrDn -Action "Set" -Status "Planned" -Note "Resolved manager DN"
                }
            } else {
                $mgrEmail = Normalize-String $r.'Manager Email'
                $mgrName  = Normalize-String $r.'Manager'
                if (-not [string]::IsNullOrWhiteSpace($mgrEmail) -or -not [string]::IsNullOrWhiteSpace($mgrName)) {
                    Write-Log "Manager could not be uniquely resolved for $($u.SamAccountName). Skipping manager update." "WARN"
                    Add-ChangeRow -SamAccountName $u.SamAccountName -UPN $u.UserPrincipalName -Attribute 'manager' -OldValue (Normalize-String $u.manager) -NewValue "" -Action "Skip" -Status "Unresolved" -Note "Manager not uniquely resolvable"
                }
            }
        }

        $hasChanges = ($replace.Count -gt 0) -or ($clear.Count -gt 0) -or ($notes -contains "SetManager")

        if (-not $hasChanges) {
            $skipped++
            Write-Log "No changes needed for $($u.SamAccountName)."
            continue
        }

        # Apply changes
        if ($WhatIfMode) {
            $updated++
            Write-Log "WhatIfMode: would update $($u.SamAccountName) with $($replace.Count) replace(s), $($clear.Count) clear(s)." "CHANGE"
            continue
        }

        # 1) Replace / Clear
        if ($replace.Count -gt 0 -or $clear.Count -gt 0) {
            $params = @{
                Identity    = $u.DistinguishedName
                ErrorAction = 'Stop'
            }
            if ($replace.Count -gt 0) { $params['Replace'] = $replace }
            if ($clear.Count   -gt 0) { $params['Clear']   = $clear }

            Set-ADUser @params
        }

        # 2) Manager (if planned)
        if ($UpdateManager) {
            $mgrDn = Resolve-ManagerDn -Row $r
            if ($mgrDn -and $mgrDn -ne (Normalize-String $u.manager)) {
                Set-ADUser -Identity $u.DistinguishedName -Manager $mgrDn -ErrorAction Stop
            }
        }

        $updated++
        Write-Log "Updated $($u.SamAccountName)." "CHANGE"

        # Mark planned changes as Applied (best-effort)
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

