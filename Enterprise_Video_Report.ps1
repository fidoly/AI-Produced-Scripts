<#
.SYNOPSIS
    A diagnostic script to investigate SharePoint search results. It will dump all available properties for any result that is missing a critical value.

.NOTES
    Author: Gemini Assistant
    Version: 7.3 (Diagnostic Edition - Complete)
#>

#region Parameters
param(
    [Parameter(Mandatory=$true, HelpMessage="Enter the SharePoint URL Prefix (e.g., 'corplogin').")]
    [string]$SharePointPrefix,

    [Parameter(Mandatory=$true, HelpMessage="Enter the Tenant ID (GUID) or a verified domain name for authentication.")]
    [string]$AuthenticationIdentifier,

    [Parameter(Mandatory=$true, HelpMessage="The Application (client) ID from the private App Registration you created in your tenant.")]
    [string]$ClientId
)
#endregion

# --- Display Header and Get Path---
Write-Host "======================================================================" -ForegroundColor Green
Write-Host "   SharePoint Online Video Report v7.3 (Diagnostic Edition)"
Write-Host "======================================================================" -ForegroundColor Green
Write-Host ""

$isValidPath = $false
while (-not $isValidPath) {
    $OutputDirectory = Read-Host -Prompt "Please enter the full path to a folder to save the report (e.g., C:\Reports)"
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) { Write-Warning "Path cannot be empty."; continue }
    if (Test-Path -Path $OutputDirectory -PathType Container) { $isValidPath = $true } 
    else {
        Write-Warning "The path '$OutputDirectory' does not exist or is not a folder."
        $choice = Read-Host "Would you like to try and create this folder? (y/n)"
        if ($choice -eq 'y') {
            try {
                New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop | Out-Null
                Write-Host "Successfully created directory: $OutputDirectory" -ForegroundColor Green
                $isValidPath = $true
            } catch { Write-Error "Failed to create directory. Please enter a new, valid path." }
        }
    }
}

# --- Start Transcript Logging ---
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = [System.IO.Path]::Combine($OutputDirectory, "VideoReport_Log_$($SharePointPrefix)_$($timestamp).log")
Start-Transcript -Path $logFile -Append
Write-Host "Report and log files will be saved to: $OutputDirectory" -ForegroundColor Cyan
Write-Host "Log started at: $(Get-Date)"
#endregion

# --- Main Logic ---
try {
    #region Prerequisite Check
    Write-Host "`n[1/4] Checking for required module 'PnP.PowerShell'..." -ForegroundColor Cyan
    if ($null -eq (Get-Module -Name PnP.PowerShell -ListAvailable)) {
        Write-Warning "Required module 'PnP.PowerShell' is not installed."
        $choice = Read-Host "May I install it for you? (y/n)"
        if ($choice -eq 'y') {
            Write-Host "Installing..." -ForegroundColor Yellow; Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        } else { throw "User declined module installation." }
    } else { Write-Host "Module is already installed." -ForegroundColor Green }
    #endregion

    #region Connection
    Write-Host "`n[2/4] Preparing secure connection..." -ForegroundColor Cyan
    $intendedPrefix = $SharePointPrefix.ToLower().Trim()
    $tenantAdminUrl = "https://$($intendedPrefix)-admin.sharepoint.com"
    Connect-PnPOnline -Url $tenantAdminUrl -DeviceLogin -Tenant $AuthenticationIdentifier -ClientId $ClientId
    Write-Host "Successfully connected." -ForegroundColor Green
    #endregion

    #region Verification
    $connection = Get-PnPConnection
    Write-Host "`n[3/4] Verifying connection..." -ForegroundColor Cyan
    $actualPrefix = ($connection.Url.ToLower() -split '[/.]')[2] -replace '-admin$'
    if ($actualPrefix -ne $intendedPrefix) {
        throw "CONNECTION MISMATCH! Intended prefix: '$($intendedPrefix)'. Actual connected prefix: '$($actualPrefix)'."
    }
    Write-Host "Connection verified. Starting search..." -ForegroundColor Green
    #endregion

    #region Reporting & Diagnostics
    Write-Host "`n[4/4] Searching for videos and generating report..." -ForegroundColor Cyan
    $videoExtensions = @("mp4", "mov", "wmv", "avi", "mpeg", "mpg", "mkv", "flv", "webm")
    $fileTypeQuery = ($videoExtensions | ForEach-Object { "FileExtension:$_" }) -join " OR "
    $searchQuery = "($fileTypeQuery) AND IsDocument:true"
    
    # Requesting many more properties for diagnostics
    $selectProperties = @(
        'Title','Path','SPSiteUrl','FileExtension','FileType','Created','LastModifiedTime',
        'Author','Size','ParentLink','SiteID','WebId','ListID','OriginalPath','ContentClass',
        'SPWebUrl','UniqueId','IsContainer'
    )

    $searchResults = Submit-PnPSearchQuery -Query $searchQuery -All -SelectProperties $selectProperties
    
    if ($null -eq $searchResults -or $searchResults.Count -eq 0) {
        Write-Warning "No video files were found in this tenant's search index with the current query."
    }
    else {
        Write-Host "Found $($searchResults.Count) video file(s). Compiling data..." -ForegroundColor Green
        $siteGeoCache = @{}
        $reportData = [System.Collections.Generic.List[psobject]]::new()
        
        foreach ($result in $searchResults) {
            $siteUrl = $result.SPSiteUrl
            
            # Dump all properties if SPSiteUrl is missing
            if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                Write-Warning "Skipping an item because its SiteURL was missing. DUMPING ALL AVAILABLE PROPERTIES FOR DIAGNOSTICS:"
                # This will print every piece of data we received for this specific result
                $result | Format-List
                continue # Skip to the next file
            }

            if (-not $siteGeoCache.ContainsKey($siteUrl)) {
                $site = Get-PnPTenantSite -Url $siteUrl -ErrorAction SilentlyContinue
                $siteGeoCache[$siteUrl] = if ($site) { $site.GeoLocation } else { "N/A" }
            }
            
            $reportData.Add([PSCustomObject]@{
                VideoName = $result.Title; VideoURL = $result.Path; SiteName = $result.SiteName
                SiteGeoLocation = $siteGeoCache[$siteUrl]; SiteURL = $siteUrl
                VideoCreatorName = $result.Author; VideoCreationDate = $result.Created
                VideoModificationDate = $result.LastModifiedTime
                VideoSizeMB = if ($result.Size -gt 0) { [math]::Round(($result.Size / 1MB), 2) } else { 0 }
            })
        }
        
        if ($reportData.Count -gt 0) {
            $outputFile = [System.IO.Path]::Combine($OutputDirectory, "VideoReport_Diagnostic_$($SharePointPrefix)_$($timestamp).csv")
            $reportData | Export-Csv -Path $outputFile -NoTypeInformation
            Write-Host "`n======================================================================" -ForegroundColor Green
            Write-Host "                        SUCCESS! REPORT COMPLETE"
            Write-Host "======================================================================" -ForegroundColor Green
            Write-Host "$($reportData.Count) valid video records have been exported to:" -ForegroundColor White
            Write-Host $outputFile
        } else {
            Write-Warning "`nReport was not generated as no processable video records were found."
        }
    }
    #endregion
}
catch {
    Write-Error "A critical error occurred. Error: $($_.Exception.Message)"
}
finally {
    # Final, robust disconnect logic
    if (Get-PnPConnection) {
        Write-Host "Disconnecting SharePoint session..." -ForegroundColor Cyan
        Disconnect-PnPOnline
    }
    Write-Host "`nScript finished at $(Get-Date)."
    Write-Host "A detailed log was saved to: $logFile" -ForegroundColor Yellow
    Stop-Transcript
}