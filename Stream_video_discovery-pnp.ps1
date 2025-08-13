<#
.SYNOPSIS
    Uses PnP PowerShell to connect to a multi-geo Microsoft 365 tenant, find all video files in all
    SharePoint sites using the search engine, and export the details to a CSV file.
.DESCRIPTION
    This script uses the robust PnP PowerShell module to reliably catalog video assets. By leveraging
    the SharePoint search engine, it avoids common permission errors with locked libraries and is
    significantly more efficient than iterating through folders.

    Features:
    - **Uses PnP PowerShell:** A proven and powerful tool for SharePoint administration.
    - **Reliable Site Discovery:** Uses Get-PnPTenantSite to discover all sites across all geos.
    - **Efficient Search-Based Discovery:** Uses Submit-PnPSearchQuery to find videos quickly.
    - **Multi-Geo Aware:** Reports the Geo-Location for each site.
    - **Robust Error Handling:** The search method naturally avoids issues with locked system libraries.
.NOTES
    Version: 5.0 (PnP Edition)
    Author: Gemini AI - Complete rewrite using PnP PowerShell for reliability and efficiency.
#>

# --- Main Script ---
try {
    # 1. Module Check for PnP.PowerShell
    Write-Host "Step 1: Checking for PnP.PowerShell module..." -ForegroundColor Yellow
    if (-not (Get-Module -Name "PnP.PowerShell" -ListAvailable)) {
        Write-Warning "PnP.PowerShell module not found."
        Write-Host "Please install it by running: Install-Module PnP.PowerShell" -ForegroundColor Cyan
        $install = Read-Host "Would you like to try and install it now? (Y/N)"
        if ($install -eq 'Y') {
            try {
                Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
                Write-Host "PnP.PowerShell installed successfully." -ForegroundColor Green
            } catch {
                Write-Error "Failed to install PnP.PowerShell. Please install it manually and re-run the script."
                return
            }
        } else {
            Write-Error "PnP.PowerShell is required. Aborting script."
            return
        }
    } else {
        Write-Host "  [OK] PnP.PowerShell is installed." -ForegroundColor Green
    }

    # 2. User Input
    Write-Host "`nStep 2: Gathering information..." -ForegroundColor Yellow
    $adminUrl = Read-Host -Prompt "Please enter your SharePoint Admin Center URL (e.g., https://yourtenant-admin.sharepoint.com)"
    if ([string]::IsNullOrWhiteSpace($adminUrl)) {
        Write-Error "The SharePoint Admin Center URL is required."
        return
    }

    $reportFolder = Read-Host -Prompt "Please enter the full path for the CSV report (e.g., C:\Reports)"
    if ([string]::IsNullOrWhiteSpace($reportFolder)) { $reportFolder = "C:\Reports" }
    if (-not (Test-Path -Path $reportFolder)) {
        Write-Host "Creating report folder: $reportFolder"
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
    }

    # 3. Connection
    Write-Host "`nStep 3: Connecting to SharePoint Online..." -ForegroundColor Yellow
    try {
        Connect-PnPOnline -Url $adminUrl -Interactive -ErrorAction Stop
        Write-Host "  [OK] Successfully connected to $($adminUrl.Split('/')[2])" -ForegroundColor Green
    } catch {
        Write-Error "Failed to connect. Please check the URL and your permissions."
        return
    }

    # 4. Data Retrieval
    Write-Host "`nStep 4: Discovering all SharePoint sites... (This is multi-geo aware)" -ForegroundColor Yellow
    $sites = Get-PnPTenantSite -ErrorAction Stop
    Write-Host "  [OK] Found $($sites.Count) sites to scan across all geographic locations." -ForegroundColor Green

    if ($sites.Count -eq 0) {
        Write-Warning "No SharePoint sites were found."
        return
    }
    
    $allVideosData = @()
    $videoExtensions = @('mp4', 'mov', 'avi', 'wmv', 'mkv', 'webm', 'm4v', 'mpg', 'mpeg')
    $scanStartTime = Get-Date

    # Build the search query
    $fileTypeClauses = $videoExtensions | ForEach-Object { "filetype:$_" }
    $searchQuery = $fileTypeClauses -join " OR "

    Write-Host "`nStep 5: Searching for video files across all sites. This may take a while..." -ForegroundColor Yellow
    $progress = 0
    foreach ($site in $sites) {
        $progress++
        Write-Progress -Activity "Searching for Videos" -Status "Querying Site: $($site.Title)" -PercentComplete (($progress / $sites.Count) * 100)
        
        try {
            # We don't need to connect to each site, we can use the search query to target it
            $siteSpecificQuery = "$searchQuery path:$($site.Url)"
            Write-Host "  Scanning Site: $($site.Title) (Geo: $($site.GeoLocation))" -ForegroundColor Cyan
            
            # Submit the search query to find all videos in the current site
            $results = Submit-PnPSearchQuery -Query $siteSpecificQuery -All -TrimDuplicates:$false -SelectProperties 'Title,Path,Author,Created,LastModifiedTime,SiteName,OriginalPath,FileSize' -ErrorAction Stop
            
            if ($results.ResultRows.Count -gt 0) {
                Write-Host "    Found $($results.ResultRows.Count) video(s)." -ForegroundColor Green

                foreach ($row in $results.ResultRows) {
                    $videoData = [PSCustomObject]@{
                        VideoName             = $row.Title
                        VideoURL              = $row.Path
                        SiteName              = $row.SiteName
                        SiteGeoLocation       = $site.GeoLocation
                        SiteURL               = $site.Url
                        VideoCreatorName      = $row.Author
                        VideoCreationDate     = $row.Created
                        VideoModificationDate = $row.LastModifiedTime
                        VideoSizeMB           = if ($row.FileSize) { [math]::Round([long]$row.FileSize / 1MB, 2) } else { 0 }
                    }
                    $allVideosData += $videoData
                }
            }
        } catch {
            Write-Warning "    Could not search site $($site.Title): $($_.Exception.Message)"
        }
    }

    # 6. Export to CSV
    $totalScanTime = [math]::Round(((Get-Date) - $scanStartTime).TotalMinutes, 1)
    Write-Host "`n" + ("-"*50)
    Write-Host "Scan Summary:" -ForegroundColor Yellow
    Write-Host "  - Total scan time: $totalScanTime minutes"
    Write-Host "  - Total sites scanned: $($sites.Count)"
    Write-Host "  - Total videos found: $($allVideosData.Count)" -ForegroundColor Green

    Write-Host "`nStep 6: Generating Report..." -ForegroundColor Yellow
    if ($allVideosData.Count -gt 0) {
        $datestring = (Get-Date).ToString("yyyyMMdd-HHmm")
        $csvfileName = Join-Path -Path $reportFolder -ChildPath "PnP_VideoReport_$datestring.csv"
        
        $allVideosData | Export-Csv -Path $csvfileName -NoTypeInformation -Encoding UTF8
        Write-Host "`nReport successfully generated!" -ForegroundColor Green
        Write-Host "Report saved to: $csvfileName"
    } else {
        Write-Warning "No video files were found across any sites."
    }

} catch {
    Write-Error "A critical, unrecoverable error occurred: $($_.Exception.Message)"
} finally {
    if (Get-PnPConnection) {
        Write-Host "`nDisconnecting from SharePoint Online..." -ForegroundColor Yellow
        Disconnect-PnPOnline
        Write-Host "  [OK] Disconnected successfully." -ForegroundColor Green
    }
    Write-Host "`nScript execution completed." -ForegroundColor Cyan
}