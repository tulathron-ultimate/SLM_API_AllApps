<#
.SYNOPSIS
    Export Snow License Manager computer + application usage data to CSV.

.DESCRIPTION
    Pulls:
      - Computers: /customers/{customerId}/computers
      - Applications per computer: /customers/{customerId}/computers/{computerId}/applications

    Joins them into a flat CSV with:
      ComputerId, ComputerName, Domain, Org, Status, LastScanDate, IsVirtual
      ApplicationId, ApplicationName, Manufacturer, Family, Bundle info
      InstallDate, DiscoveredDate, FirstUsed, LastUsed, Run, AvgUsageTime, Users
      LicenseRequired, IsInstalled, IsBlacklisted, IsWhitelisted, IsVirtual, IsOEM,
      IsMSDN, IsWebApplication, ApplicationItemCost

    "Last user" lookups are skipped by default to reduce API calls; enable with -IncludeUserLookups.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = "./SnowApiConfig.json",

    [Parameter(Mandatory = $false, HelpMessage = "Include inactive computers in the export.")]
    [switch]$IncludeInactiveComputers,

    [Parameter(Mandatory = $false, HelpMessage = "Fetch MostRecent/MostFrequent users per computer (slower).")]
    [switch]$IncludeUserLookups,

    [Parameter(Mandatory = $false, HelpMessage = "How many rows to buffer before flushing to CSV (memory control).")]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$FlushBatchSize = 1000
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found at '$ConfigPath'."
}

Write-Verbose "Loading config from $ConfigPath"
$configJson = Get-Content -LiteralPath $ConfigPath -Raw
$config     = $configJson | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($config.ApiBaseUrl)) { throw "ApiBaseUrl missing from config." }
if ([string]::IsNullOrWhiteSpace($config.CustomerId)) { throw "CustomerId missing from config." }
if ([string]::IsNullOrWhiteSpace($config.Username))   { throw "Username missing from config." }
if ([string]::IsNullOrWhiteSpace($config.Password))   { throw "Password missing from config." }

if ([string]::IsNullOrWhiteSpace($config.OutputCsvPath)) {
    $config.OutputCsvPath = "./Snow_ComputerApplicationUsage.csv"
}

# Normalize base URL
$baseUrl = $config.ApiBaseUrl.TrimEnd('/')

Write-Verbose "Using base URL: $baseUrl"
Write-Verbose "CustomerId: $($config.CustomerId)"
Write-Verbose "Output CSV: $($config.OutputCsvPath)"

function New-BasicAuthHeader {
    param(
        [Parameter(Mandatory)]
        [string]$Username,
        [Parameter(Mandatory)]
        [string]$Password
    )

    $pair   = '{0}:{1}' -f $Username, $Password
    $bytes  = [System.Text.Encoding]::UTF8.GetBytes($pair)
    $token  = [Convert]::ToBase64String($bytes)

    return @{
        'Authorization' = "Basic $token"
        'Accept'        = 'application/json'
    }
}

$headers = New-BasicAuthHeader -Username $config.Username -Password $config.Password

function Invoke-SnowGet {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    $uri =
        if ($Url -match '^https?://') {
            $Url
        }
        else {
            if ($Url.StartsWith('/')) {
                "$baseUrl$Url"
            }
            else {
                "$baseUrl/$Url"
            }
        }

    Write-Verbose "GET $uri"

    try {
        return Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
    }
    catch {
        Write-Host "Request failed for URL: $uri" -ForegroundColor Red

        $response = $_.Exception.Response
        if ($response -and $response.GetResponseStream) {
            try {
                $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
                $body   = $reader.ReadToEnd()
                Write-Host "Response body from server:" -ForegroundColor Red
                Write-Host $body
            }
            catch {
                Write-Host "Unable to read response body." -ForegroundColor Red
            }
        }

        throw
    }
}

# For collection resources that have Meta/Links/Body (Body = array of sub-resources)
function Get-SnowCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    $results = @()
    $nextUrl = $Url

    do {
        $resp = Invoke-SnowGet -Url $nextUrl

        if ($resp.Body) {
            $results += $resp.Body
        }

        # Look for a "Next" link for paging
        $nextLink = $resp.Links | Where-Object { $_.Rel -eq 'Next' } | Select-Object -First 1
        if ($nextLink) {
            $nextUrl = $nextLink.Href
            Write-Verbose "Paging: found Next link -> $($nextUrl)"
        }
        else {
            $nextUrl = $null
        }
    } while ($nextUrl)

    return $results
}

function Get-UserFromLink {
    param(
        [Parameter(Mandatory)]
        $Link,
        [Parameter(Mandatory)]
        [string]$Label
    )

    if (-not $Link -or [string]::IsNullOrWhiteSpace($Link.Href)) {
        return $null
    }

    if ($Link.Href -match '/users/0/?$') {
        Write-Verbose "Skipping $Label user lookup because link points to user id 0 (no data)."
        return $null
    }

    try {
        $uResp = Invoke-SnowGet -Url $Link.Href
        if ($uResp.Body) {
            return [pscustomobject]@{
                Id       = $uResp.Body.Id
                Username = $uResp.Body.Username
            }
        }
    }
    catch {
        Write-Warning "Failed to fetch $Label user at $($Link.Href): $($_.Exception.Message)"
    }

    return $null
}

# Optional helper: get MostRecentUser / MostFrequentUser for a computer
function Get-ComputerUserSummary {
    param(
        [Parameter(Mandatory)]
        $ComputerResource  # one element of the Computer collection (has Links + Body)
    )

    $summary = [ordered]@{
        MostRecentUserId        = $null
        MostRecentUserName      = $null
        MostFrequentUserId      = $null
        MostFrequentUserName    = $null
    }

    $mostRecentLink = $ComputerResource.Links | Where-Object { $_.Rel -eq 'MostRecentUser' } | Select-Object -First 1
    $recentUser     = Get-UserFromLink -Link $mostRecentLink -Label 'most recent'
    if ($recentUser) {
        $summary.MostRecentUserId   = $recentUser.Id
        $summary.MostRecentUserName = $recentUser.Username
    }

    $mostFrequentLink = $ComputerResource.Links | Where-Object { $_.Rel -eq 'MostFrequentUser' } | Select-Object -First 1
    $frequentUser     = Get-UserFromLink -Link $mostFrequentLink -Label 'most frequent'
    if ($frequentUser) {
        $summary.MostFrequentUserId   = $frequentUser.Id
        $summary.MostFrequentUserName = $frequentUser.Username
    }

    return [pscustomobject]$summary
}

# MAIN
Write-Host "Fetching computers from Snow..." -ForegroundColor Cyan

# IMPORTANT: no trailing slash before the query; we let the API default to JSON
$computers = Get-SnowCollection -Url "customers/$($config.CustomerId)/computers"

$filteredComputers =
    if ($IncludeInactiveComputers) {
        $computers
    }
    else {
        $computers | Where-Object { $_.Body -and $_.Body.Status -eq 'Active' }
    }

if ($computers.Count -ne $filteredComputers.Count) {
    Write-Host "Filtered out $($computers.Count - $filteredComputers.Count) non-active computers." -ForegroundColor Gray
}

Write-Host "Fetching application catalog from Snow for caching..." -ForegroundColor Cyan
$applicationCache = @{}
$allApplications  = Get-SnowCollection -Url "customers/$($config.CustomerId)/applications"
foreach ($appResource in $allApplications) {
    if ($appResource.Body -and $appResource.Body.Id) {
        $applicationCache[$appResource.Body.Id] = $appResource.Body
    }
}
Write-Host "Cached $($applicationCache.Count) applications for lookups." -ForegroundColor Gray

if (-not $filteredComputers) {
    Write-Warning "No computers returned from Snow."
    return
}

$rows = New-Object System.Collections.Generic.List[object]

# Reset CSV if it already exists so batch appends don't retain stale data
if (Test-Path -LiteralPath $config.OutputCsvPath) {
    Remove-Item -LiteralPath $config.OutputCsvPath
}

$csvInitialized = $false

foreach ($compResource in $filteredComputers) {
    $comp = $compResource.Body
    if (-not $comp) { continue }

    Write-Host "Processing computer $($comp.Id) - $($comp.Name)..." -ForegroundColor Yellow

    $userSummary =
        if ($IncludeUserLookups) {
            Get-ComputerUserSummary -ComputerResource $compResource
        }
        else {
            # Skipping user lookups to reduce API calls
            [pscustomobject]@{
                MostRecentUserId     = $null
                MostRecentUserName   = $null
                MostFrequentUserId   = $null
                MostFrequentUserName = $null
            }
        }

    # Per-computer application usage
    $apps = Get-SnowCollection -Url "customers/$($config.CustomerId)/computers/$($comp.Id)/applications"

    foreach ($appResource in $apps) {
        $app = $appResource.Body
        if (-not $app) { continue }

        $appDetails = $applicationCache[$app.Id]

        $row = [pscustomobject]@{
            # Computer context
            ComputerId                    = $comp.Id
            ComputerName                  = $comp.Name
            ComputerDomain                = $comp.Domain
            ComputerOrganization          = $comp.Organization
            ComputerStatus                = $comp.Status
            ComputerIsVirtual             = $comp.IsVirtual
            ComputerLastScanDate          = $comp.LastScanDate

            # Approximate "last user" at computer level
            MostRecentComputerUserId      = $userSummary.MostRecentUserId
            MostRecentComputerUserName    = $userSummary.MostRecentUserName
            MostFrequentComputerUserId    = $userSummary.MostFrequentUserId
            MostFrequentComputerUserName  = $userSummary.MostFrequentUserName

            # Application identity
            ApplicationId                 = $app.Id
            ApplicationName               = if ($appDetails) { $appDetails.Name } else { $app.Name }
            ApplicationManufacturerId     = if ($appDetails) { $appDetails.ManufacturerId } else { $app.ManufacturerId }
            ApplicationManufacturer       = if ($appDetails) { $appDetails.ManufacturerName } else { $app.ManufacturerName }
            ApplicationFamilyId           = if ($appDetails) { $appDetails.FamilyId } else { $app.FamilyId }
            ApplicationFamilyName         = if ($appDetails) { $appDetails.FamilyName } else { $app.FamilyName }
            BundleApplicationId           = if ($appDetails) { $appDetails.BundleApplicationId } else { $app.BundleApplicationId }
            BundleApplicationName         = if ($appDetails) { $appDetails.BundleApplicationName } else { $app.BundleApplicationName }

            # Usage & lifecycle
            FirstUsed                     = $app.FirstUsed
            LastUsed                      = $app.LastUsed
            InstallDate                   = $app.InstallDate
            DiscoveredDate                = $app.DiscoveredDate
            RunCount                      = $app.Run
            AvgUsageTimeMinutes           = $app.AvgUsageTime
            UserCount                     = $app.Users

            # Licensing / SAM flags
            LicenseRequired               = $app.LicenseRequired
            AppIsInstalled                = $app.IsInstalled
            AppIsBlacklisted              = $app.IsBlacklisted
            AppIsWhitelisted              = $app.IsWhitelisted
            AppIsVirtual                  = $app.IsVirtual
            AppIsOEM                      = $app.IsOEM
            AppIsMSDN                     = $app.IsMSDN
            AppIsWebApplication           = $app.IsWebApplication
            ApplicationItemCost           = $app.ApplicationItemCost
        }

        $rows.Add($row)

        if ($rows.Count -ge $FlushBatchSize) {
            $writeParams = @{ Path = $config.OutputCsvPath; NoTypeInformation = $true; Encoding = 'UTF8' }
            if ($csvInitialized) {
                $rows | Export-Csv @writeParams -Append
            }
            else {
                $rows | Export-Csv @writeParams
                $csvInitialized = $true
            }

            $rows.Clear()
        }
    }
}

if ($rows.Count -gt 0) {
    $writeParams = @{ Path = $config.OutputCsvPath; NoTypeInformation = $true; Encoding = 'UTF8' }
    if ($csvInitialized) {
        $rows | Export-Csv @writeParams -Append
    }
    else {
        $rows | Export-Csv @writeParams
    }
}

Write-Host "Wrote output to $($config.OutputCsvPath)." -ForegroundColor Green

Write-Host "Done." -ForegroundColor Green

<#
NEXT STEP IDEAS:

- Filter to LicenseRequired = $true and LastUsed older than X days
- Add an extra CSV built from:
    /customers/{customerId}/applications/{appId}/users
    /customers/{customerId}/users/{userId}/applications
  to get true per-user LastUsed per application.
#>
