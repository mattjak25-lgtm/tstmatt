#Requires -Version 5.1
<#
.SYNOPSIS
    IOMS eHub -> SharePoint Migration (All-in-One)
    No Python required. Runs on built-in Windows PowerShell 5.1+.

.DESCRIPTION
    Step 1 (Parse):   Reads all .htm files from a local/network folder,
                      extracts metadata and screen codes, writes CSV + JSON.

    Step 2 (Migrate): Reads the CSV, connects to SharePoint Online,
                      creates pages from template, sets all metadata.

    Run modes:
      -ParseOnly      Just parse HTML → produce CSV/JSON, no SharePoint
      -MigrateOnly    Skip parse, use existing CSV → create SP pages
      -WhatIf         Dry run migrate (shows what would happen, no changes)
      -StartRow N     Resume migration from row N (after interruption)

.PARAMETER SourceFolder
    Path to folder containing .htm files and images.
    Can be a local path or UNC network share.
    Example: "\\qcsdevweb03\IOMSeHub"

.PARAMETER SiteUrl
    SharePoint Online site URL.
    Default: https://correctionsqld.sharepoint.com/sites/QCSeHubS1

.PARAMETER CsvPath
    Output/input path for the page manifest CSV.
    Default: .\migration_manifest.csv

.PARAMETER ImageManifest
    Output path for the image metadata JSON.
    Default: .\image_manifest.json

.PARAMETER ParseOnly
    Only parse HTML files, skip SharePoint connection.

.PARAMETER MigrateOnly
    Skip parsing, use existing CsvPath for migration.

.PARAMETER WhatIf
    Dry run — show what would be created, make no changes.

.PARAMETER StartRow
    Resume migration from this row number. Default: 1

.EXAMPLE
    # Step 1: Parse only (test your HTML folder first)
    .\migrate_ioms.ps1 -SourceFolder "\\qcsdevweb03\IOMSeHub" -ParseOnly

    # Step 2: Review migration_manifest.csv in Excel, then migrate
    .\migrate_ioms.ps1 -MigrateOnly -WhatIf
    .\migrate_ioms.ps1 -MigrateOnly

    # Both steps in one go
    .\migrate_ioms.ps1 -SourceFolder "\\qcsdevweb03\IOMSeHub"

    # Resume after interruption at row 47
    .\migrate_ioms.ps1 -MigrateOnly -StartRow 47

.NOTES
    PnP.PowerShell is required for migration steps.
    Install without admin rights:
        Install-Module PnP.PowerShell -Scope CurrentUser -Force
    If PSGallery is blocked, use -ParseOnly to produce the CSV,
    then use Power Automate to create pages (see comments at end of file).
#>

param(
    [string]$SourceFolder   = "\\qcsdevweb03\IOMSeHub",
    [string]$SiteUrl        = "https://correctionsqld.sharepoint.com/sites/QCSeHubS1",
    [string]$CsvPath        = ".\migration_manifest.csv",
    [string]$ImageManifest  = ".\image_manifest.json",
    [switch]$ParseOnly,
    [switch]$MigrateOnly,
    [switch]$WhatIf,
    [int]$StartRow          = 1
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ═══════════════════════════════════════════════════════════════════════════════
# SHARED HELPERS
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Status([string]$msg, [string]$colour = "Cyan") {
    Write-Host $msg -ForegroundColor $colour
}

function Invoke-WithRetry([scriptblock]$action, [int]$maxRetries = 4) {
    $delay = 2
    for ($i = 0; $i -le $maxRetries; $i++) {
        try { return & $action }
        catch {
            if ($i -eq $maxRetries) { throw }
            Write-Status "  Retrying in ${delay}s..." "Yellow"
            Start-Sleep -Seconds $delay
            $delay *= 2
        }
    }
}

function Get-SafePageName([string]$title) {
    $safe = $title -replace '[^A-Za-z0-9\s\-]', '' `
                   -replace '\s+', '-' `
                   -replace '\-+', '-'
    return $safe.Trim('-').ToLower()
}

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — PARSE HTML FILES
# ═══════════════════════════════════════════════════════════════════════════════

function Get-HiddenInput([string]$html, [string]$id) {
    if ($html -match "id=""$id""\s+[^/]*value=""([^""]*)""|value=""([^""]*)""\s+[^/]*id=""$id""") {
        return ($Matches[1] + $Matches[2]).Trim()
    }
    return ''
}

function Get-ScreenCodes([string]$text) {
    # Matches IOMS window codes like (CSP440), (CP110), (VR100), (CSP3310)
    $matches = [regex]::Matches($text, '\(([A-Z]{2,6}\d{2,4})\)')
    $codes = @{}
    foreach ($m in $matches) { $codes[$m.Groups[1].Value] = $true }
    return ($codes.Keys | Sort-Object) -join '|'
}

function Get-BusinessArea([string]$html) {
    # Breadcrumb link in the aboveheading related topics table
    if ($html -match 'class="relatedtopics aboveheading"[\s\S]*?<a\s[^>]*>([^<]+)</a>') {
        return $Matches[1].Trim() -replace '&amp;', '&'
    }
    return ''
}

function Get-SubArea([string]$html) {
    # <p class="notepp"> contains sub-area like "Probation & Parole"
    if ($html -match '<p class="notepp">([^<]+)</p>') {
        return $Matches[1].Trim() -replace '&amp;', '&'
    }
    return ''
}

function Get-PageType([string]$html) {
    if ($html -match '<p class="subheading2">Practice Steps</p>') { return 'Procedure' }
    if ($html -match '<p class="subheading2">Overview</p>')       { return 'Overview' }
    if ($html -match '<p class="subheading2">Report</p>')         { return 'Report' }
    if ($html -match '<ol class="listnumber">')                   { return 'Procedure' }
    return 'Reference'
}

function Get-PageImages([string]$html, [string]$pageKey, [string]$pageTitle) {
    $images = @()
    $imgMatches = [regex]::Matches($html, '<img\s+id="f?(\d+)"\s+[^>]*src="([^"]+)"[^>]*/>')
    foreach ($m in $imgMatches) {
        $imgId  = $m.Groups[1].Value
        $imgSrc = $m.Groups[2].Value

        # Find the list item or paragraph containing this image
        $imgPos    = $m.Index
        $blockStart = [Math]::Max(0, $imgPos - 800)
        $context   = $html.Substring($blockStart, [Math]::Min(800, $html.Length - $blockStart))
        $nearbyCodes = Get-ScreenCodes $context

        $images += [pscustomobject]@{
            PageKey     = $pageKey
            PageTitle   = $pageTitle
            ImageSrc    = $imgSrc
            ImageID     = $imgId
            ScreenCodes = $nearbyCodes
        }
    }
    return $images
}

function Invoke-ParseFolder([string]$folder, [string]$csvOut, [string]$imgOut) {

    if (-not (Test-Path $folder)) {
        Write-Error "Source folder not found: $folder"
        exit 1
    }

    $htmFiles = @(Get-ChildItem -Path $folder -Filter "*.htm") +
                @(Get-ChildItem -Path $folder -Filter "*.html")
    $htmFiles = $htmFiles | Sort-Object Name

    if ($htmFiles.Count -eq 0) {
        Write-Error "No .htm/.html files found in: $folder"
        exit 1
    }

    Write-Status "Found $($htmFiles.Count) HTML files in $folder`n"

    $pages     = [System.Collections.Generic.List[object]]::new()
    $allImages = [System.Collections.Generic.List[object]]::new()
    $errors    = [System.Collections.Generic.List[object]]::new()

    foreach ($file in $htmFiles) {
        try {
            $html = [System.IO.File]::ReadAllText($file.FullName, [System.Text.Encoding]::UTF8)

            # Core metadata from hidden inputs
            $topicId   = Get-HiddenInput $html 'topicId'
            if (-not $topicId) { $topicId = [System.IO.Path]::GetFileNameWithoutExtension($file.Name) }

            $title = Get-HiddenInput $html 'topicDescription'
            if (-not $title) {
                if ($html -match '<title>\s*([^<]+)\s*</title>') { $title = $Matches[1].Trim() }
                else { $title = $topicId }
            }

            $modifiedRaw  = Get-HiddenInput $html 'footer-modified'
            $modifiedByRaw = Get-HiddenInput $html 'footer-modifiedby'
            $modified     = $modifiedRaw  -replace '^Last modified:\s*', ''
            $modifiedBy   = $modifiedByRaw -replace '^Modified by:\s*', ''

            $legacyFile  = "#${topicId}.htm"
            $bizArea     = Get-BusinessArea $html
            $subArea     = Get-SubArea $html
            $pagetype    = Get-PageType $html
            $screenCodes = Get-ScreenCodes $html
            $images      = Get-PageImages $html $topicId $title

            $pages.Add([pscustomobject]@{
                Title        = $title
                PageKey      = $topicId
                LegacyFile   = $legacyFile
                BusinessArea = $bizArea
                SubArea      = $subArea
                ContentType  = 'Topic'
                Pagetype     = $pagetype
                ScreenCodes  = $screenCodes
                ImageCount   = $images.Count
                LastModified = $modified
                ModifiedBy   = $modifiedBy
                Status       = 'To Do'
                SourceFile   = $file.Name
            })

            foreach ($img in $images) { $allImages.Add($img) }

            $codesDisplay = if ($screenCodes) { $screenCodes } else { 'none' }
            Write-Status ("  OK  {0,-22} '{1}'" -f $file.Name, $title.Substring(0, [Math]::Min(50,$title.Length))) "Gray"
            if ($screenCodes) {
                Write-Status ("       codes=[{0}]" -f $screenCodes) "DarkGray"
            }

        } catch {
            $errors.Add([pscustomobject]@{ File = $file.Name; Error = $_.Exception.Message })
            Write-Status "  ERR $($file.Name): $($_.Exception.Message)" "Red"
        }
    }

    # Write CSV (UTF-8 with BOM so Excel opens it correctly)
    $pages | Export-Csv -Path $csvOut -NoTypeInformation -Encoding UTF8
    Write-Status "`nPage manifest : $csvOut  ($($pages.Count) pages)" "Green"

    # Write image manifest JSON
    $allImages | ConvertTo-Json -Depth 3 | Set-Content -Path $imgOut -Encoding UTF8
    Write-Status "Image manifest: $imgOut  ($($allImages.Count) images)" "Green"

    if ($errors.Count -gt 0) {
        Write-Status "`nErrors ($($errors.Count)):" "Yellow"
        $errors | ForEach-Object { Write-Status "  $($_.File): $($_.Error)" "Yellow" }
    }

    Write-Status "`nParse complete. Review $csvOut in Excel before running migration." "White"
    return $pages.Count
}

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — SHAREPOINT MIGRATION
# ═══════════════════════════════════════════════════════════════════════════════

$sitePagesLib       = "Site Pages"
$templateRelUrl     = "SitePages/Templates/Page-template.aspx"
$defaultContentType = "Topic"

function Initialize-SiteColumns {
    Write-Status "`nChecking/creating site columns..."

    # ScreenCodes on Site Pages (pipe-separated codes, e.g. "CSP440|CSP310")
    $f1 = Get-PnPField -List $sitePagesLib -Identity "ScreenCodes" -ErrorAction SilentlyContinue
    if (-not $f1) {
        Add-PnPField -List $sitePagesLib -DisplayName "ScreenCodes" `
                     -InternalName "ScreenCodes" -Type Note -AddToDefaultView | Out-Null
        Write-Status "  Created: ScreenCodes on '$sitePagesLib'" "Green"
    } else {
        Write-Status "  Exists : ScreenCodes on '$sitePagesLib'" "Gray"
    }

    # ScreenCode on Site Assets (single image tag)
    $f2 = Get-PnPField -List "Site Assets" -Identity "ScreenCode" -ErrorAction SilentlyContinue
    if (-not $f2) {
        Add-PnPField -List "Site Assets" -DisplayName "ScreenCode" `
                     -InternalName "ScreenCode" -Type Text -AddToDefaultView | Out-Null
        Write-Status "  Created: ScreenCode on 'Site Assets'" "Green"
    } else {
        Write-Status "  Exists : ScreenCode on 'Site Assets'" "Gray"
    }
}

function Invoke-Migration([string]$csvPath, [string]$imgManifest,
                          [string]$imageFolder, [switch]$whatIf, [int]$startRow) {

    if (-not (Test-Path $csvPath)) {
        Write-Error "CSV not found: $csvPath — run with -ParseOnly first, or provide -CsvPath"
        exit 1
    }

    # Check PnP module
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Status @"

PnP.PowerShell is not installed. Install it without admin rights by running:

    Install-Module PnP.PowerShell -Scope CurrentUser -Force

If PSGallery is blocked on your network, see the Power Automate fallback
instructions at the bottom of this script.
"@ "Yellow"
        exit 1
    }

    Import-Module PnP.PowerShell -ErrorAction Stop

    Write-Status "Connecting to $SiteUrl"
    if (-not $whatIf) {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        Initialize-SiteColumns
    }

    # Optional: upload images
    if ($imageFolder -and (Test-Path $imageFolder)) {
        Invoke-ImageUpload $imageFolder $imgManifest $whatIf
    }

    $rows = Import-Csv -Path $csvPath -Encoding UTF8
    Write-Status "Loaded $($rows.Count) pages from $csvPath`n"

    $results   = [System.Collections.Generic.List[object]]::new()
    $succeeded = 0; $skipped = 0; $failed = 0; $rowNum = 0

    foreach ($row in $rows) {
        $rowNum++
        if ($rowNum -lt $startRow) { continue }

        $title       = $row.Title.Trim()
        $pageKey     = $row.PageKey.Trim()
        $legacyFile  = $row.LegacyFile.Trim()
        $bizArea     = $row.BusinessArea.Trim()
        $pagetype    = $row.Pagetype.Trim()
        $screenCodes = $row.ScreenCodes.Trim()
        $pageName    = Get-SafePageName $title
        $pageUrl     = "SitePages/$pageName.aspx"

        Write-Status "`n[$rowNum/$($rows.Count)] $title" "White"
        Write-Status "  key=$pageKey  pagetype=$pagetype  codes=$screenCodes"

        if ($whatIf) {
            Write-Status "  [WHATIF] Would create: $pageUrl" "Yellow"
            $results.Add([pscustomobject]@{
                Row = $rowNum; Title = $title; PageKey = $pageKey
                PageUrl = $pageUrl; Result = "WhatIf"
            })
            continue
        }

        try {
            # Skip if already exists
            $existing = Invoke-WithRetry {
                Get-PnPFile -Url $pageUrl -ErrorAction SilentlyContinue
            }
            if ($existing) {
                Write-Status "  Skipping — already exists" "Yellow"
                $skipped++
                $results.Add([pscustomobject]@{
                    Row = $rowNum; Title = $title; PageKey = $pageKey
                    PageUrl = $pageUrl; Result = "Skipped"
                })
                continue
            }

            # 1. Copy template
            Invoke-WithRetry {
                Copy-PnPFile -SourceUrl $templateRelUrl -TargetUrl $pageUrl -Force | Out-Null
            }
            Write-Status "  Copied template" "Green"

            # 2. Set title
            Set-PnPPage -Identity $pageName -Title $title | Out-Null

            # 3. Get page list item ID
            $pnpPage = Invoke-WithRetry { Get-PnPPage -Identity $pageName }

            # 4. Set metadata
            # NOTE: If "ContentType" conflicts with SP reserved field, change to "ContentType0"
            #       Check List Settings -> Content Type column internal name if this errors.
            Invoke-WithRetry {
                Set-PnPListItem -List $sitePagesLib -Identity $pnpPage.PageId -Values @{
                    "PageKey"      = $pageKey
                    "LegacyFile"   = $legacyFile
                    "BusinessArea" = $bizArea
                    "ContentType"  = $defaultContentType
                    "Pagetype"     = $pagetype
                    "ScreenCodes"  = $screenCodes
                    "Status"       = "To Do"
                } | Out-Null
            }
            Write-Status "  Metadata set" "Green"

            $succeeded++
            $results.Add([pscustomobject]@{
                Row = $rowNum; Title = $title; PageKey = $pageKey
                PageUrl = $pageUrl; Result = "Created"
            })

        } catch {
            $err = $_.Exception.Message
            Write-Status "  FAILED: $err" "Red"
            $failed++
            $results.Add([pscustomobject]@{
                Row = $rowNum; Title = $title; PageKey = $pageKey
                PageUrl = $pageUrl; Result = "Failed: $err"
            })
        }

        Start-Sleep -Milliseconds 500   # avoid throttling
    }

    $logPath = "migration_run_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $results | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

    Write-Status "`n──────────────────────────────────────────" "White"
    Write-Status "Created: $succeeded   Skipped: $skipped   Failed: $failed" "White"
    Write-Status "Run log: $logPath" "White"
    if ($whatIf) { Write-Status "`nDry run only. Remove -WhatIf to execute." "Yellow" }
}

function Invoke-ImageUpload([string]$folder, [string]$manifest, [switch]$whatIf) {
    Write-Status "`nUploading images to SiteAssets/ioms-legacy..."

    $imgData = @{}
    if (Test-Path $manifest) {
        $raw = Get-Content $manifest -Raw | ConvertFrom-Json
        foreach ($img in $raw) { $imgData[$img.ImageSrc] = $img.ScreenCodes }
    }

    if (-not $whatIf) {
        Invoke-WithRetry { Resolve-PnPFolder -SiteRelativePath "SiteAssets/ioms-legacy" | Out-Null }
    }

    $files = Get-ChildItem -Path $folder -Include "*.jpg","*.png","*.gif" -File
    $count = 0
    foreach ($f in $files) {
        $codes = $imgData[$f.Name]
        Write-Status "  $($f.Name)  codes=[$codes]" "Gray"
        if (-not $whatIf) {
            Invoke-WithRetry {
                $up = Add-PnPFile -Path $f.FullName -Folder "SiteAssets/ioms-legacy"
                if ($codes) {
                    Set-PnPListItem -List "Site Assets" `
                        -Identity $up.ListItemAllFields.Id `
                        -Values @{ "ScreenCode" = $codes } | Out-Null
                }
            }
        }
        $count++
    }
    Write-Status "Uploaded $count images." "Green"
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════

if (-not $MigrateOnly) {
    Write-Status "═══ STEP 1: PARSING HTML FILES ═══" "White"
    Invoke-ParseFolder $SourceFolder $CsvPath $ImageManifest
}

if (-not $ParseOnly) {
    Write-Status "`n═══ STEP 2: SHAREPOINT MIGRATION ═══" "White"
    Invoke-Migration $CsvPath $ImageManifest "" $WhatIf $StartRow
}

<#
──────────────────────────────────────────────────────────────────────────────
POWER AUTOMATE FALLBACK (if PnP.PowerShell install is blocked)
──────────────────────────────────────────────────────────────────────────────
1. Run parse step only:
       .\migrate_ioms.ps1 -SourceFolder "\\qcsdevweb03\IOMSeHub" -ParseOnly

2. Open migration_manifest.csv in Excel Online → Table → Export to SP List
   (or manually import via SharePoint "Quick Edit" grid view)

3. In Power Automate (make.powerautomate.com), create a flow:
   Trigger : Manually trigger a flow
   Action 1: Get items — from your migration list
   Action 2: Apply to each item
     Action 2a: Send an HTTP request to SharePoint
       Method : POST
       Uri    : _api/web/getfolderbyserverrelativeurl('SitePages')/Files/
                AddTemplateFile(urlOfFile='SitePages/@{item()?['Title']}.aspx',templateType=1)
       Headers: { "Accept": "application/json;odata=verbose",
                  "Content-Type": "application/json;odata=verbose",
                  "X-RequestDigest": "<form digest>" }
     Action 2b: Send an HTTP request to SharePoint (update metadata)
       Method : POST/MERGE on the list item to set PageKey, LegacyFile, etc.

   See: https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-rest-endpoints
──────────────────────────────────────────────────────────────────────────────
#>
