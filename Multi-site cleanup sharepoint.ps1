# Requires:
# - PnP.PowerShell module
# - A configured Entra ID application/client ID suitable for interactive PnP authentication
# - Appropriate permissions to every SharePoint site and library listed below
#
# IMPORTANT:
# $ApplyChanges = $false performs a dry run.
# $ApplyChanges = $true removes the selected historical versions.
#
# $PermanentlyDeleteRemovedVersions = $false
#     Removed versions are sent to the SharePoint recycle bin.
#     Change this to $true only after reviewing the proposed actions and confirming that you want to permanently delete the selected historical versions.
#
# $PermanentlyDeleteRemovedVersions = $true
#     Only the exact historical versions selected by this script are permanently
#     deleted. They are not sent to the recycle bin.
#
# The script DOES NOT empty the site's existing recycle bin.
# The script only removes the historical versions selected by this script. 
#
# Sites are processed sequentially.
# A failure on one site does not prevent remaining sites from being processed.
# Reauthentication is performed for each site.
#
# Created by Kim Rune Lundberg. Copilot helped generate some of the code, 
# but I reviewed and modified it to ensure it met my requirements. Script has also been tested and verified to work as intended.



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
$ClientId = "00000000-0000-0000-0000-000000000000" #Add in your Entra ID application/client ID here.


# File extensions that will be processed.
$FileExtensions = @(
    ".dwg",
    ".dxf",
    ".rvt",
    ".ifc",
    ".zip",
    ".tmp",
    ".lock",
    ".pln",
    ".tm",
    ".psd",
    ".jpg",
    ".tiff",
    ".pdf"
)


# Total number of versions to retain INCLUDING the current version.
#
# 1  = current version only
# 2  = current + newest historical version
# 10 = current + 9 newest historical versions
$KeepTotal = 1


# =============================================================================
# SAFETY SWITCHES
# =============================================================================

# Default to a dry run; set to $true only after reviewing the proposed actions.
$ApplyChanges = $false


# $false = send removed versions to SharePoint recycle bin
# $true  = permanently delete the exact versions selected by this script
#
# IMPORTANT:
# This does NOT empty the existing recycle bin.
# When live changes are enabled, recycle selected versions by default rather
# than permanently deleting them.
$PermanentlyDeleteRemovedVersions = $false


# =============================================================================
# CONFIGURATION VALIDATION
# =============================================================================

if (-not $SiteUrls -or $SiteUrls.Count -eq 0) {
    throw '$SiteUrls must contain at least one SharePoint site URL.'
}


foreach ($configuredSiteUrl in $SiteUrls) {

    if ([string]::IsNullOrWhiteSpace($configuredSiteUrl)) {
        throw '$SiteUrls contains an empty site URL.'
    }
}


# Remove duplicate site URLs.
$SiteUrls = @(
    $SiteUrls |
        Select-Object -Unique
)


if ([string]::IsNullOrWhiteSpace($LibraryName)) {
    throw '$LibraryName cannot be empty.'
}


if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw '$ClientId cannot be empty.'
}


if ($KeepTotal -lt 1) {
    throw '$KeepTotal must be at least 1 because the current version is always retained.'
}


if (-not $FileExtensions -or $FileExtensions.Count -eq 0) {
    throw '$FileExtensions must contain at least one file extension.'
}


if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {

    throw @"
PnP.PowerShell is not installed.

Install it with:

Install-Module PnP.PowerShell -Scope CurrentUser
"@
}


# =============================================================================
# LOAD PNP.POWERSHELL
# =============================================================================

Import-Module PnP.PowerShell -ErrorAction Stop


$loadedPnPModule = Get-Module PnP.PowerShell |
    Sort-Object Version -Descending |
    Select-Object -First 1


if ($loadedPnPModule) {

    Write-Host (
        "Using PnP.PowerShell {0} from {1}" -f
        $loadedPnPModule.Version,
        $loadedPnPModule.Path
    ) -ForegroundColor DarkGray
}


# =============================================================================
# NORMALIZE FILE EXTENSIONS
# =============================================================================

$FileExtensions = @(
    $FileExtensions |
        ForEach-Object {

            $extension = $_.Trim().ToLowerInvariant()

            if (-not $extension.StartsWith(".")) {
                $extension = ".$extension"
            }

            $extension
        } |
        Sort-Object -Unique
)


# Get-PnPFileVersion returns historical versions only.
#
# The current version is not included.
#
# Therefore:
#
# KeepTotal = 1
# keepPrevious = 0
#
# KeepTotal = 10
# keepPrevious = 9

$keepPrevious = $KeepTotal - 1


# =============================================================================
# GLOBAL COUNTERS
# =============================================================================

$totalSites             = $SiteUrls.Count
$sitesProcessed         = 0
$sitesPartial           = 0
$sitesFailed            = 0

$totalProcessedFiles    = 0
$totalSkippedFiles      = 0

$totalCandidateVersions = 0
$totalRemovedVersions   = 0

$totalRecycledVersions  = 0
$totalPermanentVersions = 0

$totalFailedVersions    = 0
$totalFailedFiles       = 0


# Per-site results used for the final table.
$siteResults = @()


# =============================================================================
# PROCESS EACH SITE
# =============================================================================

foreach ($SiteUrl in $SiteUrls) {


    # =========================================================================
    # PER-SITE COUNTERS
    # =========================================================================

    $processedFiles    = 0
    $skippedFiles      = 0

    $candidateVersions = 0
    $removedVersions   = 0

    $recycledVersions  = 0
    $permanentVersions = 0

    $failedVersions    = 0
    $failedFiles       = 0

    $siteSucceeded     = $false
    $siteFailureReason = $null

    $connection        = $null


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

        # =====================================================================
        # CONNECT TO SITE
        # =====================================================================

        $connection = Connect-PnPOnline `
            -Url $SiteUrl `
            -Interactive `
            -ClientId $ClientId `
            -ReturnConnection `
            -ErrorAction Stop


        if ($null -eq $connection) {
            throw "Connect-PnPOnline did not return a connection for '$SiteUrl'."
        }


        Write-Host "Connected successfully." -ForegroundColor Green
        Write-Host ""


        # =====================================================================
        # ENUMERATE DOCUMENT LIBRARY
        # =====================================================================

        Get-PnPListItem `
            -List $LibraryName `
            -PageSize 500 `
            -Fields "FileLeafRef", "FileRef", "FSObjType" `
            -Connection $connection `
            -ErrorAction Stop |
        ForEach-Object {

            $item = $_


            # -----------------------------------------------------------------
            # SKIP FOLDERS
            # -----------------------------------------------------------------

            # FSObjType:
            # 0 = file
            # 1 = folder

            if ([int]$item.FieldValues["FSObjType"] -ne 0) {
                return
            }


            # -----------------------------------------------------------------
            # GET FILE NAME
            # -----------------------------------------------------------------

            $name = [string]$item.FieldValues["FileLeafRef"]


            if ([string]::IsNullOrWhiteSpace($name)) {

                $skippedFiles++

                return
            }


            # -----------------------------------------------------------------
            # CHECK FILE EXTENSION
            # -----------------------------------------------------------------

            $extension = [System.IO.Path]::GetExtension($name)


            if ([string]::IsNullOrWhiteSpace($extension)) {

                $skippedFiles++

                return
            }


            $extension = $extension.ToLowerInvariant()


            if ($extension -notin $FileExtensions) {

                $skippedFiles++

                return
            }


            # -----------------------------------------------------------------
            # GET SERVER-RELATIVE FILE URL
            # -----------------------------------------------------------------

            $fileUrl = [string]$item.FieldValues["FileRef"]


            if ([string]::IsNullOrWhiteSpace($fileUrl)) {

                Write-Warning (
                    "[{0}] Skipping '{1}' because its FileRef value is empty." -f
                    $SiteUrl,
                    $name
                )

                $skippedFiles++

                return
            }


            $processedFiles++


            # =================================================================
            # PROCESS FILE VERSIONS
            # =================================================================

            try {

                # Get historical versions.
                #
                # The current version is NOT included in this collection.

                $versions = @(
                    Get-PnPFileVersion `
                        -Url $fileUrl `
                        -Connection $connection `
                        -ErrorAction Stop |
                    Sort-Object -Property Created -Descending
                )


                $totalVersionCount = $versions.Count + 1


                # -------------------------------------------------------------
                # FILE ALREADY WITHIN VERSION LIMIT
                # -------------------------------------------------------------

                if ($versions.Count -le $keepPrevious) {

                    Write-Host (
                        "KEEP: {0} ({1} total version(s))" -f
                        $fileUrl,
                        $totalVersionCount
                    ) -ForegroundColor DarkGreen

                    return
                }


                # -------------------------------------------------------------
                # SELECT VERSIONS FOR REMOVAL
                # -------------------------------------------------------------

                # Versions are sorted:
                #
                # newest
                #   |
                #   v
                # oldest
                #
                # Skip the historical versions we want to keep.

                $versionsToRemove = @(
                    $versions |
                        Select-Object -Skip $keepPrevious
                )


                # -------------------------------------------------------------
                # REMOVE VERSIONS
                # -------------------------------------------------------------

                foreach ($version in $versionsToRemove) {

                    $candidateVersions++


                    # ---------------------------------------------------------
                    # VERIFY VERSION ID
                    # ---------------------------------------------------------

                    if ($null -eq $version.Id) {

                        $failedVersions++

                        Write-Warning (
                            "[{0}] Cannot remove version {1} from '{2}' because its version ID is null." -f
                            $SiteUrl,
                            $version.VersionLabel,
                            $fileUrl
                        )

                        continue
                    }


                    # ---------------------------------------------------------
                    # DRY RUN
                    # ---------------------------------------------------------

                    if (-not $ApplyChanges) {

                        $plannedAction = if ($PermanentlyDeleteRemovedVersions) {
                            "permanently delete"
                        }
                        else {
                            "recycle"
                        }


                        Write-Host (
                            "DRY RUN: Would {0} {1} | version {2} | ID {3} | created {4}" -f
                            $plannedAction,
                            $fileUrl,
                            $version.VersionLabel,
                            $version.Id,
                            $version.Created
                        ) -ForegroundColor Cyan


                        continue
                    }


                    # ---------------------------------------------------------
                    # LIVE REMOVAL
                    # ---------------------------------------------------------

                    try {

                        # Build parameters for the exact historical version.
                        #
                        # If permanent deletion is FALSE:
                        #     -Recycle is added.
                        #
                        # If permanent deletion is TRUE:
                     #
                        # This means only the version selected by Identity is
                        # permanently removed. Existing recycle-bin entries are
                        # not touched.

                        $removeVersionParameters = @{
                            Url         = $fileUrl
                            Identity    = $version.Id
                            Force       = $true
                            Connection  = $connection
                            ErrorAction = "Stop"
                        }


                        if (-not $PermanentlyDeleteRemovedVersions) {
                            $removeVersionParameters.Recycle = $true
                        }


                        Remove-PnPFileVersion @removeVersionParameters


                        $removedVersions++


                        # -----------------------------------------------------
                        # UPDATE ACTION COUNTERS
                        # -----------------------------------------------------

                        if ($PermanentlyDeleteRemovedVersions) {

                            $permanentVersions++

                            $actionText  = "PERMANENTLY DELETED"
                            $actionColor = "Red"
                        }
                        else {

                            $recycledVersions++

                            $actionText  = "RECYCLED"
                            $actionColor = "Yellow"
                        }


                        # -----------------------------------------------------
                        # REPORT SUCCESS
                        # -----------------------------------------------------

                        Write-Host (
                            "{0}: {1} | version {2} | ID {3} | created {4}" -f
                            $actionText,
                            $fileUrl,
                            $version.VersionLabel,
                            $version.Id,
                            $version.Created
                        ) -ForegroundColor $actionColor

                    }
                    catch {

                        $failedVersions++


                        Write-Warning (
                            "[{0}] Failed to remove version {1} (ID {2}) from '{3}': {4}" -f
                            $SiteUrl,
                            $version.VersionLabel,
                            $version.Id,
                            $fileUrl,
                            $_.Exception.Message
                        )


                        Write-Host (
                            "  Exception type: {0}" -f
                            $_.Exception.GetType().FullName
                        ) -ForegroundColor DarkRed


                        Write-Host (
                            "  FullyQualifiedErrorId: {0}" -f
                            $_.FullyQualifiedErrorId
                        ) -ForegroundColor DarkRed


                        Write-Host (
                            "  CategoryInfo: {0}" -f
                            $_.CategoryInfo
                        ) -ForegroundColor DarkRed
                    }
                }
            }
            catch {

                $failedFiles++


                Write-Warning (
                    "[{0}] Failed to retrieve or process versions for '{1}': {2}" -f
                    $SiteUrl,
                    $fileUrl,
                    $_.Exception.Message
                )
            }
        }


        # =====================================================================
        # SITE COMPLETED
        # =====================================================================

        $siteSucceeded = $true
    }
    catch {

        # =====================================================================
        # SITE FAILURE
        # =====================================================================

        $siteFailureReason = $_.Exception.Message


        Write-Warning (
            "Site processing failed for '{0}': {1}" -f
            $SiteUrl,
            $siteFailureReason
        )
    }
    finally {

        # =====================================================================
        # DISCONNECT SITE
        # =====================================================================

        if ($null -ne $connection) {

            try {

                Disconnect-PnPOnline `
                    -Connection $connection `
                    -ErrorAction SilentlyContinue

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


    # This status distinguishes fully successful sites from sites where the
    # library was processed but one or more files or versions failed.
    $siteStatus = if (-not $siteSucceeded) {
        "Failed"
    }
    elseif ($failedVersions -gt 0 -or $failedFiles -gt 0) {
        "PartialFailure"
    }
    else {
        "Completed"
    }

    switch ($siteStatus) {
        "Completed" {
            $sitesProcessed++
        }
        "PartialFailure" {
            $sitesPartial++
            $siteFailureReason = "$failedFiles file(s) and $failedVersions version(s) failed."
        }
        "Failed" {
            $sitesFailed++
        }
    }


    # =========================================================================
    # UPDATE GLOBAL COUNTERS
    # =========================================================================

    $totalProcessedFiles    += $processedFiles
    $totalSkippedFiles      += $skippedFiles

    $totalCandidateVersions += $candidateVersions
    $totalRemovedVersions   += $removedVersions

    $totalRecycledVersions  += $recycledVersions
    $totalPermanentVersions += $permanentVersions

    $totalFailedVersions    += $failedVersions
    $totalFailedFiles       += $failedFiles


    # =========================================================================
    # STORE SITE RESULT
    # =========================================================================

    $siteResults += [PSCustomObject]@{

        SiteUrl            = $SiteUrl

        Status             = $siteStatus

        MatchingFiles      = $processedFiles
        SkippedFiles       = $skippedFiles

        VersionsIdentified = $candidateVersions
        VersionsRemoved    = $removedVersions

        VersionsRecycled   = $recycledVersions
        PermanentDeletes   = $permanentVersions

        VersionFailures    = $failedVersions
        FileFailures       = $failedFiles

        FailureReason      = $siteFailureReason
    }


    # =========================================================================
    # SITE SUMMARY
    # =========================================================================

    Write-Host ""
    Write-Host "Site summary: $SiteUrl" -ForegroundColor Green
    Write-Host "------------------------------------------------------------"

    Write-Host "Matching files processed : $processedFiles"
    Write-Host "Nonmatching files skipped : $skippedFiles"

    Write-Host "Versions identified       : $candidateVersions"
    Write-Host "Versions removed          : $removedVersions"

    Write-Host "Versions recycled         : $recycledVersions"
    Write-Host "Permanent deletions       : $permanentVersions"

    Write-Host "Version removal failures  : $failedVersions"
    Write-Host "File processing failures  : $failedFiles"

    Write-Host (
        "Status                    : {0}" -f
        $siteStatus
    )


    if ($siteFailureReason) {

        Write-Host (
            "Failure reason            : {0}" -f
            $siteFailureReason
        ) -ForegroundColor Red
    }
}


# =============================================================================
# FINAL SUMMARY
# =============================================================================

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "FINAL PROCESSING SUMMARY" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green

Write-Host "Sites configured            : $totalSites"
Write-Host "Sites completed successfully: $sitesProcessed"
Write-Host "Sites partially failed      : $sitesPartial"
Write-Host "Sites failed                : $sitesFailed"

Write-Host "Matching files processed    : $totalProcessedFiles"
Write-Host "Nonmatching files skipped   : $totalSkippedFiles"

Write-Host "Versions identified         : $totalCandidateVersions"
Write-Host "Versions removed            : $totalRemovedVersions"

Write-Host "Versions recycled           : $totalRecycledVersions"
Write-Host "Permanent deletions         : $totalPermanentVersions"

Write-Host "Version removal failures    : $totalFailedVersions"
Write-Host "File processing failures    : $totalFailedFiles"

Write-Host "Keep total                  : $KeepTotal"
Write-Host "Apply changes               : $ApplyChanges"
Write-Host "Permanent version deletion  : $PermanentlyDeleteRemovedVersions"

Write-Host ""


# =============================================================================
# PER-SITE SUMMARY TABLE
# =============================================================================

$siteResults |
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


# =============================================================================
# FINAL MODE MESSAGE
# =============================================================================

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


# Signal unsuccessful or partial processing to schedulers and other automation.
if ($sitesFailed -gt 0 -or $sitesPartial -gt 0) {
    throw "Cleanup finished with failures: $sitesFailed site(s) failed, $sitesPartial site(s) had file or version failures."
}