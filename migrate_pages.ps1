#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
    IOMS eHub -> SharePoint Migration Script
    Creates SharePoint pages from template and sets metadata from CSV manifest.

.DESCRIPTION
    Reads migration_manifest.csv produced by parse_ioms.py and:
      1. Copies the page template to create each new page
      2. Sets all metadata columns (PageKey, LegacyFile, BusinessArea, etc.)
      3. Optionally uploads images to Site Assets with screen code metadata

.PARAMETER SiteUrl
    SharePoint site URL. Default: https://correctionsqld.sharepoint.com/sites/QCSeHubS1

.PARAMETER CsvPath
    Path to migration_manifest.csv. Default: .\migration_manifest.csv

.PARAMETER ImageFolder
    Local folder containing source images. If provided, uploads images to Site Assets.

.PARAMETER ImageManifest
    Path to image_manifest.json. Used to tag images with screen codes.

.PARAMETER WhatIf
    Dry run — show what would be created without making changes.

.PARAMETER StartRow
    Skip to a specific row number (for resuming interrupted runs). Default: 1

.EXAMPLE
    # Dry run first
    .\migrate_pages.ps1 -WhatIf

    # Run for real
    .\migrate_pages.ps1

    # Resume from row 50
    .\migrate_pages.ps1 -StartRow 50

    # Include image upload
    .\migrate_pages.ps1 -ImageFolder "C:\ioms_images" -ImageManifest ".\image_manifest.json"
#>

param(
    [string]$SiteUrl      = "https://correctionsqld.sharepoint.com/sites/QCSeHubS1",
    [string]$CsvPath      = ".\migration_manifest.csv",
    [string]$ImageFolder  = "",
    [string]$ImageManifest = ".\image_manifest.json",
    [switch]$WhatIf,
    [int]$StartRow        = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Helpers ──────────────────────────────────────────────────────────────────

function Write-Status([string]$msg, [string]$colour = "Cyan") {
    Write-Host $msg -ForegroundColor $colour
}

function Invoke-WithRetry([scriptblock]$action, [int]$maxRetries = 4) {
    $delay = 2
    for ($i = 0; $i -le $maxRetries; $i++) {
        try {
            return & $action
        } catch {
            if ($i -eq $maxRetries) { throw }
            Write-Status "  Throttled/error, retrying in ${delay}s..." "Yellow"
            Start-Sleep -Seconds $delay
            $delay *= 2
        }
    }
}

function Sanitize-PageName([string]$title) {
    # Convert title to a safe SharePoint page filename
    $safe = $title -replace '[^\w\s-]', '' `
                   -replace '\s+', '-' `
                   -replace '-+', '-'
    return $safe.Trim('-').ToLower()
}

# ── Business area -> badge colour mapping ─────────────────────────────────────
# Extend this table to match your actual badge colours
$BadgeColours = @{
    "Custodial"                 = "#107C10"   # green
    "Community Corrections"     = "#0078D4"   # blue
    "Community Service & Work"  = "#0078D4"   # blue
    "Specialist Operations"     = "#8764B8"   # purple
    "Probation & Parole"        = "#D83B01"   # orange
}

function Get-BadgeColour([string]$area) {
    foreach ($key in $BadgeColours.Keys) {
        if ($area -like "*$key*") { return $BadgeColours[$key] }
    }
    return "#605E5C"  # default grey
}

function Initialize-SiteColumns {
    <#
    .SYNOPSIS
        One-time setup: creates ScreenCodes column on Site Pages and Site Assets
        if they don't already exist. Safe to re-run — skips existing columns.
    #>
    Write-Status "`nInitializing site columns..."

    # ScreenCodes on Site Pages library (pipe-separated, e.g. "CSP440|CSP310")
    $spField = Get-PnPField -List $sitePagesLib -Identity "ScreenCodes" -ErrorAction SilentlyContinue
    if (-not $spField) {
        Add-PnPField -List $sitePagesLib `
                     -DisplayName "ScreenCodes" `
                     -InternalName "ScreenCodes" `
                     -Type Note `
                     -AddToDefaultView | Out-Null
        Write-Status "  Created ScreenCodes column on '$sitePagesLib'" "Green"
    } else {
        Write-Status "  ScreenCodes already exists on '$sitePagesLib'" "Gray"
    }

    # ScreenCode on Site Assets library (for image tagging)
    $saField = Get-PnPField -List "Site Assets" -Identity "ScreenCode" -ErrorAction SilentlyContinue
    if (-not $saField) {
        Add-PnPField -List "Site Assets" `
                     -DisplayName "ScreenCode" `
                     -InternalName "ScreenCode" `
                     -Type Text `
                     -AddToDefaultView | Out-Null
        Write-Status "  Created ScreenCode column on 'Site Assets'" "Green"
    } else {
        Write-Status "  ScreenCode already exists on 'Site Assets'" "Gray"
    }
}

# ── Connect ───────────────────────────────────────────────────────────────────

Write-Status "Connecting to $SiteUrl"
if (-not $WhatIf) {
    # Interactive login — works on closed government tenants
    Connect-PnPOnline -Url $SiteUrl -Interactive
}

# ── Load CSV ──────────────────────────────────────────────────────────────────

if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV not found: $CsvPath  — run parse_ioms.py first."
    exit 1
}

$pages = Import-Csv -Path $CsvPath -Encoding UTF8
Write-Status "Loaded $($pages.Count) pages from $CsvPath"

$templateRelUrl  = "SitePages/Templates/Page-template.aspx"
$sitePagesLib    = "Site Pages"
$defaultContentType = "Topic"

# ── One-time column setup ─────────────────────────────────────────────────────

if (-not $WhatIf) {
    Initialize-SiteColumns
}

# ── Image upload (optional) ───────────────────────────────────────────────────

if ($ImageFolder -and (Test-Path $ImageFolder)) {
    Write-Status "`nUploading images to Site Assets/ioms-legacy..."

    $imageData = @{}
    if (Test-Path $ImageManifest) {
        $raw = Get-Content $ImageManifest -Raw | ConvertFrom-Json
        foreach ($img in $raw) {
            $imageData[$img.ImageSrc] = $img.ScreenCodes
        }
    }

    $targetFolder = "SiteAssets/ioms-legacy"

    if (-not $WhatIf) {
        Invoke-WithRetry {
            Resolve-PnPFolder -SiteRelativePath $targetFolder | Out-Null
        }
    }

    $imageFiles = Get-ChildItem -Path $ImageFolder -Include "*.jpg","*.png","*.gif" -File
    $imgCount = 0

    foreach ($imgFile in $imageFiles) {
        $screenCodes = $imageData[$imgFile.Name]
        Write-Status "  Uploading $($imgFile.Name)  codes=[$screenCodes]" "Gray"

        if (-not $WhatIf) {
            Invoke-WithRetry {
                $uploaded = Add-PnPFile -Path $imgFile.FullName -Folder $targetFolder
                if ($screenCodes) {
                    Set-PnPListItem -List "Site Assets" -Identity $uploaded.ListItemAllFields.Id `
                        -Values @{ "ScreenCode" = $screenCodes } | Out-Null
                }
            }
        }
        $imgCount++
    }
    Write-Status "Uploaded $imgCount images." "Green"
}

# ── Page creation ─────────────────────────────────────────────────────────────

$results   = [System.Collections.Generic.List[object]]::new()
$succeeded = 0
$skipped   = 0
$failed    = 0
$rowNum    = 0

foreach ($page in $pages) {
    $rowNum++
    if ($rowNum -lt $StartRow) { continue }

    $title      = $page.Title.Trim()
    $pageKey    = $page.PageKey.Trim()
    $legacyFile = $page.LegacyFile.Trim()
    $bizArea    = $page.BusinessArea.Trim()
    $contentType = $page.ContentType.Trim()
    $pagetype   = $page.Pagetype.Trim()
    $screenCodes = $page.ScreenCodes.Trim()
    $status     = "To Do"

    $pageName   = Sanitize-PageName $title
    $pageUrl    = "SitePages/$pageName.aspx"

    Write-Status "`n[$rowNum/$($pages.Count)] $title" "White"
    Write-Status "  PageKey=$pageKey  File=$legacyFile  Name=$pageName"

    if ($WhatIf) {
        Write-Status "  [WHATIF] Would create $pageUrl" "Yellow"
        $results.Add([pscustomobject]@{
            Row        = $rowNum
            Title      = $title
            PageKey    = $pageKey
            PageUrl    = $pageUrl
            Result     = "WhatIf"
        })
        continue
    }

    try {
        # Check if page already exists (resume support)
        $existing = Invoke-WithRetry {
            Get-PnPFile -Url $pageUrl -ErrorAction SilentlyContinue
        }

        if ($existing) {
            Write-Status "  Already exists — skipping (use -StartRow or delete to recreate)" "Yellow"
            $skipped++
            $results.Add([pscustomobject]@{
                Row     = $rowNum
                Title   = $title
                PageKey = $pageKey
                PageUrl = $pageUrl
                Result  = "Skipped"
            })
            continue
        }

        # 1. Copy template to new page
        Invoke-WithRetry {
            Copy-PnPFile -SourceUrl $templateRelUrl `
                         -TargetUrl $pageUrl `
                         -Force `
                         -OverwriteIfAlreadyExists:$false | Out-Null
        }
        Write-Status "  Created from template" "Green"

        # 2. Set page title (updates both the file and the Title field)
        $pnpPage = Invoke-WithRetry {
            Get-PnPPage -Identity $pageName
        }
        Set-PnPPage -Identity $pageName -Title $title | Out-Null

        # 3. Set metadata columns on the list item
        # Note: "ContentType" is a reserved SP column name. If Set-PnPListItem
        # throws on it, check the field's actual InternalName in List Settings
        # and update the key below (commonly "ContentType0" on migrated sites).
        $metaValues = @{
            "PageKey"      = $pageKey
            "LegacyFile"   = $legacyFile
            "BusinessArea" = $bizArea
            "ContentType"  = $defaultContentType
            "Pagetype"     = $pagetype
            "ScreenCodes"  = $screenCodes   # empty string is fine if no codes
            "Status"       = $status
        }

        Invoke-WithRetry {
            Set-PnPListItem -List $sitePagesLib `
                            -Identity $pnpPage.PageId `
                            -Values $metaValues | Out-Null
        }
        Write-Status "  Metadata set: BusinessArea='$bizArea'  Pagetype='$pagetype'" "Green"

        $succeeded++
        $results.Add([pscustomobject]@{
            Row        = $rowNum
            Title      = $title
            PageKey    = $pageKey
            PageUrl    = $pageUrl
            Result     = "Created"
        })

    } catch {
        $errMsg = $_.Exception.Message
        Write-Status "  FAILED: $errMsg" "Red"
        $failed++
        $results.Add([pscustomobject]@{
            Row        = $rowNum
            Title      = $title
            PageKey    = $pageKey
            PageUrl    = $pageUrl
            Result     = "Failed: $errMsg"
        })
    }

    # Polite pause — avoids SharePoint throttling on large batches
    Start-Sleep -Milliseconds 500
}

# ── Summary ───────────────────────────────────────────────────────────────────

$logPath = "migration_run_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

Write-Status "`n─────────────────────────────────────────" "White"
Write-Status "Done.  Created: $succeeded  Skipped: $skipped  Failed: $failed" "White"
Write-Status "Run log: $logPath" "White"

if ($WhatIf) {
    Write-Status "`nThis was a dry run. Remove -WhatIf to execute." "Yellow"
}
