
# Requires:
# - PnP.PowerShell module
# - A configured Entra ID application/client ID suitable for interactive PnP authentication
# - Appropriate permissions to every SharePoint site and library listed below
#
# IMPORTANT:
# $ApplyChanges = $false performs a dry run.
# $ApplyChanges = $true removes the selected historical versions.
#
# $PermanentlyDeleteRemovedVersions = $false sends removed versions to the
# SharePoint recycle bin. Set it to $true only after reviewing the proposed
# actions and confirming permanent deletion is intended.
#
# This script does not empty any site's existing recycle bin.
# Sites are processed sequentially; failure on one site does not prevent the
# remaining sites from being processed. Reauthentication is performed per site.

# =============================================================================
# CONFIGURATION
# =============================================================================

$SiteUrls = @(
    "https://name.sharepoint.com/sites/sitename",
    "https://name.sharepoint.com/sites/sitename",
    "https://name.sharepoint.com/sites/sitename"
)

# Assumes the same library exists on every site.
$LibraryName = "Document library name"

# Entra ID application/client ID used for PnP interactive authentication.
$ClientId = "00000000-0000-0000-0000-000000000000" # Add your Entra ID application/client ID here.

$FileExtensions = @(
    # Microsoft Word
    ".doc"
    ".docx"
    ".docm"
    ".dot"
    ".dotx"
    ".dotm"


    # Microsoft Excel
    ".xls"
    ".xlsx"
    ".xlsm"
    ".xlsb"
    ".xlt"
    ".xltx"
    ".xltm"
    ".xla"
    ".xlam"


    # Microsoft PowerPoint
    ".ppt"
    ".pptx"
    ".pptm"
    ".pot"
    ".potx"
    ".potm"
    ".pps"
    ".ppsx"
    ".ppsm"
    ".ppa"
    ".ppam"
    ".sldx"
    ".sldm"
    ".thmx"


    # Microsoft Visio
    ".vsd"
    ".vsdx"
    ".vsdm"
    ".vss"
    ".vssx"
    ".vssm"
    ".vst"
    ".vstx"
    ".vstm"


    # Microsoft Project
    ".mpp"
    ".mpt"


    # Microsoft Access
    ".mdb"
    ".accdb"
    ".accde"
    ".accdr"
    ".accdt"


    # Microsoft Publisher
    ".pub"


    # Microsoft OneNote
    ".one"


    # Microsoft Outlook and email
    ".msg"
    ".eml"
    ".oft"
)

# Total number of versions to retain INCLUDING the current version.
# 1 = current version only; 10 = current + 9 newest historical versions.
$KeepTotal = 10

# Default to a dry run; set to $true only after reviewing the proposed actions.
$ApplyChanges = $false

# $false = send removed versions to SharePoint recycle bin
# $true  = permanently delete the exact versions selected by this script
$PermanentlyDeleteRemovedVersions = $false

# ===========================================================================
# Configuration validation
# ===========================================================================

if (-not $SiteUrls -or $SiteUrls.Count -eq 0) {
    throw '$SiteUrls must contain at least one SharePoint site URL.'
}

foreach ($configuredSiteUrl in $SiteUrls) {
    if ([string]::IsNullOrWhiteSpace($configuredSiteUrl)) {
        throw '$SiteUrls contains an empty site URL.'
    }
}

$SiteUrls = @($SiteUrls | Select-Object -Unique)

foreach ($configuredSiteUrl in $SiteUrls) {
    if (-not [System.Uri]::IsWellFormedUriString(
            $configuredSiteUrl,
            [System.UriKind]::Absolute
        )) {
        throw "'$configuredSiteUrl' is not a valid absolute URL."
    }
}

if ([string]::IsNullOrWhiteSpace($LibraryName)) {
    throw '$LibraryName cannot be empty.'
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw '$ClientId cannot be empty.'
}

$ParsedClientId = [guid]::Empty

if (-not [guid]::TryParse($ClientId, [ref]$ParsedClientId)) {
    throw "'$ClientId' is not a valid GUID."
}

if ($KeepTotal -lt 1) {
    throw '$KeepTotal must be at least 1 because the current version is always retained.'
}

if (-not $FileExtensions -or $FileExtensions.Count -eq 0) {
    throw '$FileExtensions must contain at least one file extension.'
}

# ===========================================================================
# Locate and import PnP.PowerShell
# ===========================================================================

$InstalledPnPModule = Get-Module -ListAvailable -Name "PnP.PowerShell" |
    Sort-Object -Property Version -Descending |
    Select-Object -First 1

if ($null -eq $InstalledPnPModule) {
    throw @"
PnP.PowerShell is not installed.

Open PowerShell 7 and run:

Install-Module PnP.PowerShell -Scope CurrentUser
"@
}

$ImportModuleParameters = @{
    Name            = "PnP.PowerShell"
    RequiredVersion = $InstalledPnPModule.Version
    ErrorAction     = "Stop"
}

Import-Module @ImportModuleParameters

$LoadedPnPModule = Get-Module -Name "PnP.PowerShell" |
    Sort-Object -Property Version -Descending |
    Select-Object -First 1

if ($null -eq $LoadedPnPModule) {
    throw 'PnP.PowerShell was found but could not be confirmed as loaded.'
}

Write-Host (
    "Using PnP.PowerShell {0} from {1}" -f
    $LoadedPnPModule.Version,
    $LoadedPnPModule.Path
) -ForegroundColor DarkGray

# ===========================================================================
# Normalize extension list
# ===========================================================================

$FileExtensions = @(
    $FileExtensions |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_)
        } |
        ForEach-Object {
            $Extension = ([string]$_).Trim().ToLowerInvariant()

            if (-not $Extension.StartsWith('.')) {
                $Extension = ".$Extension"
            }

            $Extension
        } |
        Sort-Object -Unique
)

 # Get-PnPFileVersion returns historical versions only; the current version
# is not included in this collection.
$KeepPrevious = $KeepTotal - 1

# ===========================================================================
# GLOBAL COUNTERS
# ===========================================================================

$TotalSites             = $SiteUrls.Count
$SitesProcessed         = 0
$SitesPartial           = 0
$SitesFailed            = 0
$TotalProcessedFiles    = 0
$TotalSkippedFiles      = 0
$TotalCandidateVersions = 0
$TotalRemovedVersions   = 0
$TotalRecycledVersions  = 0
$TotalPermanentVersions = 0
$TotalFailedVersions    = 0
$TotalFailedFiles       = 0
$SiteResults            = @()

# ===========================================================================
# PROCESS EACH SITE
# ===========================================================================

foreach ($SiteUrl in $SiteUrls) {
    $ProcessedFiles    = 0
    $SkippedFiles      = 0
    $CandidateVersions = 0
    $RemovedVersions   = 0
    $RecycledVersions  = 0
    $PermanentVersions = 0
    $FailedVersions    = 0
    $FailedFiles       = 0
    $SiteSucceeded     = $false
    $SiteFailureReason = $null
    $Connection        = $null

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor DarkCyan
    Write-Host "Processing site:  $SiteUrl" -ForegroundColor Cyan
    Write-Host "Library:          $LibraryName"
    Write-Host "Keep total:       $KeepTotal"
    Write-Host "Apply changes:    $ApplyChanges"
    Write-Host "Permanent delete: $PermanentlyDeleteRemovedVersions"
    Write-Host "============================================================" -ForegroundColor DarkCyan
    Write-Host ""

    try {
        $Connection = Connect-PnPOnline `
            -Url $SiteUrl `
            -Interactive `
            -ClientId $ClientId `
            -ReturnConnection `
            -ErrorAction Stop

        if ($null -eq $Connection) {
            throw "Connect-PnPOnline did not return a connection for '$SiteUrl'."
        }

        Write-Host "Connected successfully." -ForegroundColor Green
        Write-Host ""

        Get-PnPListItem `
            -List $LibraryName `
            -PageSize 500 `
            -Fields "FileLeafRef", "FileRef", "FSObjType" `
            -Connection $Connection `
            -ErrorAction Stop |
        ForEach-Object {
            $Item = $_

            if ([int]$Item.FieldValues["FSObjType"] -ne 0) {
                return
            }

            $FileName = [string]$Item.FieldValues["FileLeafRef"]
            if ([string]::IsNullOrWhiteSpace($FileName)) {
                $SkippedFiles++
                return
            }

            $Extension = [System.IO.Path]::GetExtension($FileName)
            if ([string]::IsNullOrWhiteSpace($Extension)) {
                $SkippedFiles++
                return
            }

            if ($Extension.ToLowerInvariant() -notin $FileExtensions) {
                $SkippedFiles++
                return
            }

            $FileUrl = [string]$Item.FieldValues["FileRef"]
            if ([string]::IsNullOrWhiteSpace($FileUrl)) {
                Write-Warning (
                    "[{0}] Skipping '{1}' because its FileRef value is empty." -f
                    $SiteUrl,
                    $FileName
                )
                $SkippedFiles++
                return
            }

            $ProcessedFiles++

            try {
                $Versions = @(
                    Get-PnPFileVersion `
                        -Url $FileUrl `
                        -Connection $Connection `
                        -ErrorAction Stop |
                    Sort-Object -Property Created -Descending
                )

                if ($Versions.Count -le $KeepPrevious) {
                    Write-Host (
                        "KEEP: {0} ({1} total version(s))" -f
                        $FileUrl,
                        ($Versions.Count + 1)
                    ) -ForegroundColor DarkGreen
                    return
                }

                $VersionsToRemove = @($Versions | Select-Object -Skip $KeepPrevious)

                foreach ($Version in $VersionsToRemove) {
                    $CandidateVersions++

                    if ($null -eq $Version.Id) {
                        $FailedVersions++
                        Write-Warning (
                            "[{0}] Cannot remove version {1} from '{2}' because its version ID is null." -f
                            $SiteUrl,
                            $Version.VersionLabel,
                            $FileUrl
                        )
                        continue
                    }

                    if (-not $ApplyChanges) {
                        $PlannedAction = if ($PermanentlyDeleteRemovedVersions) {
                            "permanently delete"
                        }
                        else {
                            "recycle"
                        }

                        Write-Host (
                            "DRY RUN: Would {0} {1} | version {2} | ID {3} | created {4}" -f
                            $PlannedAction,
                            $FileUrl,
                            $Version.VersionLabel,
                            $Version.Id,
                            $Version.Created
                        ) -ForegroundColor Cyan
                        continue
                    }

                    try {
                        $RemoveVersionParameters = @{
                            Url         = $FileUrl
                            Identity    = $Version.Id
                            Force       = $true
                            Connection  = $Connection
                            ErrorAction = "Stop"
                        }

                        if (-not $PermanentlyDeleteRemovedVersions) {
                            $RemoveVersionParameters.Recycle = $true
                        }

                        Remove-PnPFileVersion @RemoveVersionParameters
                        $RemovedVersions++

                        if ($PermanentlyDeleteRemovedVersions) {
                            $PermanentVersions++
                            $ActionText = "PERMANENTLY DELETED"
                            $ActionColor = "Red"
                        }
                        else {
                            $RecycledVersions++
                            $ActionText = "RECYCLED"
                            $ActionColor = "Yellow"
                        }

                        Write-Host (
                            "{0}: {1} | version {2} | ID {3} | created {4}" -f
                            $ActionText,
                            $FileUrl,
                            $Version.VersionLabel,
                            $Version.Id,
                            $Version.Created
                        ) -ForegroundColor $ActionColor
                    }
                    catch {
                        $FailedVersions++
                        Write-Warning (
                            "[{0}] Failed to remove version {1} (ID {2}) from '{3}': {4}" -f
                            $SiteUrl,
                            $Version.VersionLabel,
                            $Version.Id,
                            $FileUrl,
                            $_.Exception.Message
                        )
                        Write-Host "  Exception type: $($_.Exception.GetType().FullName)" -ForegroundColor DarkRed
                        Write-Host "  FullyQualifiedErrorId: $($_.FullyQualifiedErrorId)" -ForegroundColor DarkRed
                        Write-Host "  CategoryInfo: $($_.CategoryInfo)" -ForegroundColor DarkRed
                    }
                }
            }
            catch {
                $FailedFiles++
                Write-Warning (
                    "[{0}] Failed to retrieve or process versions for '{1}': {2}" -f
                    $SiteUrl,
                    $FileUrl,
                    $_.Exception.Message
                )
            }
        }

        $SiteSucceeded = $true
    }
    catch {
        $SiteFailureReason = $_.Exception.Message
        Write-Warning (
            "Site processing failed for '{0}': {1}" -f
            $SiteUrl,
            $SiteFailureReason
        )
    }
    finally {
        if ($null -ne $Connection) {
            try {
                Disconnect-PnPOnline -Connection $Connection -ErrorAction Stop
            }
            catch {
                Write-Warning (
                    "Could not cleanly disconnect from '{0}': {1}" -f
                    $SiteUrl,
                    $_.Exception.Message
                )
            }
        }
    }

    $SiteStatus = if (-not $SiteSucceeded) {
        "Failed"
    }
    elseif ($FailedVersions -gt 0 -or $FailedFiles -gt 0) {
        "PartialFailure"
    }
    else {
        "Completed"
    }

    switch ($SiteStatus) {
        "Completed" { $SitesProcessed++ }
        "PartialFailure" {
            $SitesPartial++
            $SiteFailureReason = "$FailedFiles file(s) and $FailedVersions version(s) failed."
        }
        "Failed" { $SitesFailed++ }
    }

    $TotalProcessedFiles    += $ProcessedFiles
    $TotalSkippedFiles      += $SkippedFiles
    $TotalCandidateVersions += $CandidateVersions
    $TotalRemovedVersions   += $RemovedVersions
    $TotalRecycledVersions  += $RecycledVersions
    $TotalPermanentVersions += $PermanentVersions
    $TotalFailedVersions    += $FailedVersions
    $TotalFailedFiles       += $FailedFiles

    $SiteResults += [PSCustomObject]@{
        SiteUrl            = $SiteUrl
        Status             = $SiteStatus
        MatchingFiles      = $ProcessedFiles
        SkippedFiles       = $SkippedFiles
        VersionsIdentified = $CandidateVersions
        VersionsRemoved    = $RemovedVersions
        VersionsRecycled   = $RecycledVersions
        PermanentDeletes   = $PermanentVersions
        VersionFailures    = $FailedVersions
        FileFailures       = $FailedFiles
        FailureReason      = $SiteFailureReason
    }

    Write-Host ""
    Write-Host "Site summary: $SiteUrl" -ForegroundColor Green
    Write-Host "------------------------------------------------------------"
    Write-Host "Matching files processed : $ProcessedFiles"
    Write-Host "Nonmatching files skipped : $SkippedFiles"
    Write-Host "Versions identified       : $CandidateVersions"
    Write-Host "Versions removed          : $RemovedVersions"
    Write-Host "Versions recycled         : $RecycledVersions"
    Write-Host "Permanent deletions       : $PermanentVersions"
    Write-Host "Version removal failures  : $FailedVersions"
    Write-Host "File processing failures  : $FailedFiles"
    Write-Host "Status                    : $SiteStatus"
    if ($SiteFailureReason) {
        Write-Host "Failure reason            : $SiteFailureReason" -ForegroundColor Red
    }
}

# ===========================================================================
# FINAL SUMMARY
# ===========================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "FINAL PROCESSING SUMMARY" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host "Sites configured            : $TotalSites"
Write-Host "Sites completed successfully: $SitesProcessed"
Write-Host "Sites partially failed      : $SitesPartial"
Write-Host "Sites failed                : $SitesFailed"
Write-Host "Matching files processed    : $TotalProcessedFiles"
Write-Host "Nonmatching files skipped   : $TotalSkippedFiles"
Write-Host "Versions identified         : $TotalCandidateVersions"
Write-Host "Versions removed            : $TotalRemovedVersions"
Write-Host "Versions recycled           : $TotalRecycledVersions"
Write-Host "Permanent deletions         : $TotalPermanentVersions"
Write-Host "Version removal failures    : $TotalFailedVersions"
Write-Host "File processing failures    : $TotalFailedFiles"
Write-Host "Keep total                  : $KeepTotal"
Write-Host "Apply changes               : $ApplyChanges"
Write-Host "Permanent version deletion  : $PermanentlyDeleteRemovedVersions"

$SiteResults |
    Select-Object `
        SiteUrl,
        Status,
        MatchingFiles,
        VersionsIdentified,
        VersionsRemoved,
        VersionsRecycled,
        PermanentDeletes,
        VersionFailures,
        FileFailures |
    Format-Table -AutoSize

if (-not $ApplyChanges) {
    Write-Host ""
    Write-Host "DRY RUN ONLY: No versions were removed." -ForegroundColor Cyan
}
elseif ($PermanentlyDeleteRemovedVersions) {
    Write-Host ""
    Write-Host (
        "LIVE RUN: Selected historical versions were permanently deleted " +
        "without using the SharePoint recycle bin."
    ) -ForegroundColor Red
}
else {
    Write-Host ""
    Write-Host (
        "LIVE RUN: Selected historical versions were sent to the " +
        "SharePoint recycle bin."
    ) -ForegroundColor Yellow
}

if ($SitesFailed -gt 0 -or $SitesPartial -gt 0) {
    throw "Cleanup finished with failures: $SitesFailed site(s) failed, $SitesPartial site(s) had file or version failures."
}