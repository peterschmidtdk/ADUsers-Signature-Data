<#
.SYNOPSIS
    Export AD user profile fields for email signature use (address/contact + ExchAttr1-15 + manager info).

.DESCRIPTION
    Exports common signature placeholder fields from on-prem AD users in a specific OU to CSV.
    Includes address fields, phone fields, webpage, and Exchange custom attributes (ExchAttr1-15),
    plus manager display name/email/title/phones.

.NOTES
    Author  : Peter
    Script  : Export-ADUsers-SignatureData.ps1
    Version : 1.4
    Updated : 2025-12-15
    Output  : Defaults to .\

.REQUIREMENTS
    - RSAT ActiveDirectory module
#>

# -----------------------------
# Config
# -----------------------------
$OU               = "OU=MyBusiness,DC=contoso,DC=local"
$OutputDirectory  = ".\"
$IncludeDisabled  = $false            # set $true to include disabled accounts
$IncludeNoEmail   = $false            # set $true to include users without mail/primary SMTP
$ExportPhotos     = $false            # set $true to export thumbnailPhoto to files (NOT into CSV)
$PhotoFolder      = Join-Path $OutputDirectory "AD_UserPhotos"

# -----------------------------
# Prep
# -----------------------------
Import-Module ActiveDirectory -ErrorAction Stop

if (-not (Test-Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
}

if ($ExportPhotos -and -not (Test-Path $PhotoFolder)) {
    New-Item -Path $PhotoFolder -ItemType Directory -Force | Out-Null
}

$Timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutFile   = Join-Path $OutputDirectory "AD_Users_SignatureData_Export_$Timestamp.csv"

# Properties used by signature placeholders (AD attributes behind them)
$props = @(
    # Identity / routing
    "SamAccountName","UserPrincipalName","DisplayName","GivenName","Surname","Initials",
    "mail","proxyAddresses",

    # Org / role
    "Company","Department","Title","Description","Info","physicalDeliveryOfficeName",

    # Address
    "streetAddress","postOfficeBox","l","st","postalCode","c","co",

    # Phones / web
    "telephoneNumber","otherTelephone","mobile","ipPhone","homePhone","facsimileTelephoneNumber","pager","wWWHomePage",

    # Manager
    "manager",

    # Exchange custom attributes (ExchAttr1-15)
    "extensionAttribute1","extensionAttribute2","extensionAttribute3","extensionAttribute4","extensionAttribute5",
    "extensionAttribute6","extensionAttribute7","extensionAttribute8","extensionAttribute9","extensionAttribute10",
    "extensionAttribute11","extensionAttribute12","extensionAttribute13","extensionAttribute14","extensionAttribute15",

    # Optional useful IDs
    "employeeID","employeeNumber","employeeType",

    # Photo presence
    "thumbnailPhoto",

    # Status
    "Enabled"
)

# Column order (also used to write header if no rows export)
$Columns = @(
    "SamAccountName","UPN","Display Name","First Name","Last Name","Initials","Email","ProxyAddresses",
    "Company","Department","Office","Position","Description","Notes",
    "Street","P.O. Box","City","State","Postal code","Country (c)","Country (co)",
    "Phone","Other phones","Mobile","IP phone","Home phone","Fax","Pager","Web page",
    "Manager","Manager Email","Manager Title","Manager Phone","Manager Mobile",
    "ExchAttr1","ExchAttr2","ExchAttr3","ExchAttr4","ExchAttr5","ExchAttr6","ExchAttr7","ExchAttr8","ExchAttr9","ExchAttr10",
    "ExchAttr11","ExchAttr12","ExchAttr13","ExchAttr14","ExchAttr15",
    "EmployeeID","EmployeeNumber","EmployeeType",
    "Enabled","HasPhoto"
)

# Cache manager lookups (performance)
$mgrCache = @{}

function Get-ManagerDetails {
    param([string]$ManagerDn)

    if ([string]::IsNullOrWhiteSpace($ManagerDn)) { return $null }

    if ($mgrCache.ContainsKey($ManagerDn)) { return $mgrCache[$ManagerDn] }

    try {
        $m = Get-ADUser -Identity $ManagerDn -Properties DisplayName,mail,Title,telephoneNumber,mobile
        $obj = [pscustomobject]@{
            DisplayName = $m.DisplayName
            Email       = $m.mail
            Title       = $m.Title
            Phone       = $m.telephoneNumber
            Mobile      = $m.mobile
        }
    }
    catch {
        $obj = [pscustomobject]@{ DisplayName=""; Email=""; Title=""; Phone=""; Mobile="" }
    }

    $mgrCache[$ManagerDn] = $obj
    return $obj
}

function Get-PrimarySmtpFromProxyAddresses {
    param([string[]]$ProxyAddresses)

    if (-not $ProxyAddresses) { return "" }

    # Primary SMTP is typically the one with uppercase "SMTP:"
    $primary = $ProxyAddresses | Where-Object { $_ -like "SMTP:*" } | Select-Object -First 1
    if ($primary) { return ($primary -replace "^SMTP:", "") }

    # Fallback: first smtp:
    $fallback = $ProxyAddresses | Where-Object { $_ -like "smtp:*" } | Select-Object -First 1
    if ($fallback) { return ($fallback -replace "^smtp:", "") }

    return ""
}

# -----------------------------
# Export
# -----------------------------
Write-Host "Loading users from OU: $OU"

$allUsers = Get-ADUser -Filter * -SearchBase $OU -Properties $props
$totalFound = @($allUsers).Count

# Count skips for visibility
$disabledCount = @($allUsers | Where-Object { $_.Enabled -ne $true }).Count
$enabledUsers  = if ($IncludeDisabled) { $allUsers } else { $allUsers | Where-Object { $_.Enabled -eq $true } }

Write-Host "Total users found        : $totalFound"
Write-Host "Disabled users in OU     : $disabledCount"
Write-Host "IncludeDisabled          : $IncludeDisabled"
Write-Host "IncludeNoEmail           : $IncludeNoEmail"

$exportList = New-Object System.Collections.Generic.List[object]

$enabledCount = @($enabledUsers).Count
$skippedNoEmail = 0

for ($i = 0; $i -lt $enabledCount; $i++) {
    $u = $enabledUsers[$i]

    $who = if ($u.DisplayName) { $u.DisplayName } else { $u.SamAccountName }
    $pct = if ($enabledCount -gt 0) { [int](($i + 1) / $enabledCount * 100) } else { 100 }

    Write-Progress -Activity "Exporting AD users" -Status "[$($i+1)/$enabledCount] $who" -PercentComplete $pct

    $mgr = Get-ManagerDetails -ManagerDn $u.manager

    # Determine Email: prefer mail, else primary SMTP from proxyAddresses
    $primarySmtp = ""
    if (-not [string]::IsNullOrWhiteSpace($u.mail)) {
        $primarySmtp = $u.mail
    } else {
        $primarySmtp = Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $u.proxyAddresses
    }

    if (-not $IncludeNoEmail -and [string]::IsNullOrWhiteSpace($primarySmtp)) {
        $skippedNoEmail++
        continue
    }

    if ($ExportPhotos -and $u.thumbnailPhoto) {
        $safeName  = ($u.SamAccountName -replace '[^a-zA-Z0-9._-]', '_')
        $photoPath = Join-Path $PhotoFolder "$safeName.jpg"
        try { [System.IO.File]::WriteAllBytes($photoPath, $u.thumbnailPhoto) } catch { }
    }

    $exportList.Add([pscustomobject]@{
        # Identity
        "SamAccountName" = $u.SamAccountName
        "UPN"            = $u.UserPrincipalName
        "Display Name"   = $u.DisplayName
        "First Name"     = $u.GivenName
        "Last Name"      = $u.Surname
        "Initials"       = $u.Initials
        "Email"          = $primarySmtp
        "ProxyAddresses" = if ($u.proxyAddresses) { ($u.proxyAddresses -join ";") } else { "" }

        # Org
        "Company"        = $u.Company
        "Department"     = $u.Department
        "Office"         = $u.physicalDeliveryOfficeName
        "Position"       = $u.Title
        "Description"    = $u.Description
        "Notes"          = $u.Info

        # Address
        "Street"         = $u.streetAddress
        "P.O. Box"       = $u.postOfficeBox
        "City"           = $u.l
        "State"          = $u.st
        "Postal code"    = $u.postalCode
        "Country (c)"    = $u.c
        "Country (co)"   = $u.co

        # Phones / web
        "Phone"          = $u.telephoneNumber
        "Other phones"   = if ($u.otherTelephone) { ($u.otherTelephone -join ";") } else { "" }
        "Mobile"         = $u.mobile
        "IP phone"       = $u.ipPhone
        "Home phone"     = $u.homePhone
        "Fax"            = $u.facsimileTelephoneNumber
        "Pager"          = $u.pager
        "Web page"       = $u.wWWHomePage

        # Manager
        "Manager"        = $mgr.DisplayName
        "Manager Email"  = $mgr.Email
        "Manager Title"  = $mgr.Title
        "Manager Phone"  = $mgr.Phone
        "Manager Mobile" = $mgr.Mobile

        # ExchAttr1-15
        "ExchAttr1"      = $u.extensionAttribute1
        "ExchAttr2"      = $u.extensionAttribute2
        "ExchAttr3"      = $u.extensionAttribute3
        "ExchAttr4"      = $u.extensionAttribute4
        "ExchAttr5"      = $u.extensionAttribute5
        "ExchAttr6"      = $u.extensionAttribute6
        "ExchAttr7"      = $u.extensionAttribute7
        "ExchAttr8"      = $u.extensionAttribute8
        "ExchAttr9"      = $u.extensionAttribute9
        "ExchAttr10"     = $u.extensionAttribute10
        "ExchAttr11"     = $u.extensionAttribute11
        "ExchAttr12"     = $u.extensionAttribute12
        "ExchAttr13"     = $u.extensionAttribute13
        "ExchAttr14"     = $u.extensionAttribute14
        "ExchAttr15"     = $u.extensionAttribute15

        # Optional IDs
        "EmployeeID"     = $u.employeeID
        "EmployeeNumber" = $u.employeeNumber
        "EmployeeType"   = $u.employeeType

        # Status
        "Enabled"        = $u.Enabled
        "HasPhoto"       = [bool]$u.thumbnailPhoto
    })
}

Write-Progress -Activity "Exporting AD users" -Completed

$exportedCount = $exportList.Count
$skippedDisabled = if ($IncludeDisabled) { 0 } else { $disabledCount }

# If nothing exported, still create a CSV with headers (so it's not empty/0 bytes)
if ($exportedCount -eq 0) {
    ($Columns -join ",") | Set-Content -Path $OutFile -Encoding UTF8
    Write-Warning "No users were exported. A header-only CSV was created."
} else {
    $exportList | Select-Object $Columns | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
}

Write-Host ""
Write-Host "Export complete: $OutFile"
Write-Host "Summary:"
Write-Host "  Total found in OU      : $totalFound"
Write-Host "  Considered (after disabled filter): $enabledCount"
Write-Host "  Exported               : $exportedCount"
Write-Host "  Skipped (disabled)     : $skippedDisabled"
Write-Host "  Skipped (no email)     : $skippedNoEmail"
if ($ExportPhotos) { Write-Host "  Photos folder          : $PhotoFolder" }
