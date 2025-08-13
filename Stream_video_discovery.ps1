<#
.SYNOPSIS
    Connects to Microsoft 365, finds all video files in SharePoint sites across all geographic locations,
    extracts their metadata, and exports the details to a CSV file.
.DESCRIPTION
    This script provides a modern, multi-geo aware solution for cataloging video assets stored in Stream (on SharePoint).
    It uses the Microsoft Graph API, which automatically handles discovering resources from all provisioned geos.

    Features:
    - **Multi-Geo Aware:** Captures the geographic location (e.g., NAM, CAN, EUR) of each site.
    - Comprehensive pre-flight check to configure the PowerShell environment.
    - Robust, looping authentication.
    - Searches all SharePoint sites for video files (.mp4, .mov, etc.).
    - Extracts detailed metadata for each video.
    - Implements robust error catching for smoother execution.
    - Exports the collected data to a timestamped CSV file.
.NOTES
    Version: 4.1
    Author: Gemini AI-Fixed version to correct the invalid 'dataLocationCode' property request.
    Creation Date: 2025-08-04
#>

# --- Main Script ---
try {
    # 0. PowerShell Version Check
    Write-Host "Checking PowerShell version..." -ForegroundColor Yellow
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        Write-Error "This script requires PowerShell 5.0 or higher. You are running PowerShell $($PSVersionTable.PSVersion)."
        return
    }
    Write-Host "  [OK] PowerShell $($PSVersionTable.PSVersion) detected." -ForegroundColor Green
    
    # 1. PowerShell Environment Pre-flight Check and Module Installation
    Write-Host "`nStep 1: Running PowerShell environment pre-flight checks..." -ForegroundColor Yellow

    # Stage A: Verify the PowerShell Gallery repository is registered and trusted.
    Write-Host "  - Checking PowerShell Gallery repository configuration..."
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Host "    [CONFIGURING] PSGallery is not trusted. Setting it to trusted to allow module installation." -ForegroundColor Cyan
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        } else {
            Write-Host "    [OK] PSGallery repository is registered and trusted." -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "    PSGallery repository not found or inaccessible. Attempting to register the default repository."
        try {
            Register-PSRepository -Default -ErrorAction Stop
            Write-Host "    [SUCCESS] PSGallery has been registered. Setting it to trusted." -ForegroundColor Green
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        }
        catch {
            Write-Error "CRITICAL: Could not find or register the PSGallery repository. Module installation is not possible."
            return
        }
    }

    # Stage B: Module Installation
    Write-Host "  - Checking for required Microsoft Graph modules..."
    
    $requiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Sites",
        "Microsoft.Graph.Files"
    )

    foreach ($module in $requiredModules) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            Write-Host "    [INSTALLING] Module '$module' not found. Installing..." -ForegroundColor Cyan
            try {
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Host "    [SUCCESS] Module '$module' has been installed." -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install module '$module': $($_.Exception.Message)"
                return
            }
        } else {
            Write-Host "    [OK] Module '$module' is already installed." -ForegroundColor Green
        }
    }

    # Stage C: Import modules
    Write-Host "  - Loading modules into the session..."
    
    try {
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Sites -Force -ErrorAction Stop
        Import-Module Microsoft.Graph.Files -Force -ErrorAction Stop
        Write-Host "    [OK] All modules loaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to import modules: $($_.Exception.Message)"
        return
    }

    Write-Host "Module pre-flight check complete." -ForegroundColor Yellow
    Write-Host ("-"*50)

    # 2. User Interface: Ask for the output folder path
    Write-Host "Step 2: Configuring output location..." -ForegroundColor Yellow
    $reportFolder = $null
    do {
        $defaultPath = "C:\Reports"
        if (-not (Test-Path -Path $defaultPath)) { $defaultPath = $env:USERPROFILE }
        $reportFolder = Read-Host -Prompt "Please enter the full path for the CSV report (e.g., $defaultPath)"

        if ([string]::IsNullOrWhiteSpace($reportFolder)) { $reportFolder = $defaultPath }

        if (-not (Test-Path -Path $reportFolder -PathType Container)) {
            Write-Warning "The path '$reportFolder' does not exist."
            if ((Read-Host -Prompt "Create it? (Y/N)") -eq 'Y') {
                try {
                    New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
                    Write-Host "Folder created." -ForegroundColor Green
                }
                catch {
                    Write-Warning "Failed to create folder: $($_.Exception.Message)"
                    $reportFolder = $null
                }
            } else { $reportFolder = $null }
        }
    } while ($null -eq $reportFolder)

    Write-Host "Report will be saved to '$reportFolder'." -ForegroundColor Green
    Write-Host ("-"*50)

    # 3. Authentication
    Write-Host "Step 3: Connecting to Microsoft Graph..." -ForegroundColor Yellow
    
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Write-Host "Disconnecting existing session..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    
    $tenantId = Read-Host -Prompt "Enter your tenant ID (optional, but recommended)"
    $scopes = @("Sites.Read.All", "Files.Read.All", "User.Read.All")
    
    try {
        $connectParams = @{ Scopes = $scopes; ErrorAction = 'Stop' }
        if (![string]::IsNullOrWhiteSpace($tenantId)) { $connectParams.TenantId = $tenantId }
        
        Write-Host "Initiating interactive login. A browser window will open." -ForegroundColor Cyan
        Connect-MgGraph @connectParams
        
        $context = Get-MgContext
        Write-Host "`nSuccessfully connected!" -ForegroundColor Green
        Write-Host "  Account: $($context.Account)"
        Write-Host "  Tenant: $($context.TenantId)"
    }
    catch {
        Write-Error "Authentication failed: $($_.Exception.Message). Script cannot continue."
        return
    }
    Write-Host ("-"*50)
    
    # 4. Data Retrieval
    Write-Host "Step 4: Retrieving data from SharePoint Online..." -ForegroundColor Yellow
    
    # Get all sites across all Geos
    Write-Host "Fetching all SharePoint sites... (This is multi-geo aware)"
    $sites = @()
    try {
        # --- CORRECTED CODE ---
        # Removed the invalid -Property parameter. The API returns a rich object by default.
        $allSites = Get-MgSite -All -ErrorAction Stop
    }
    catch {
        Write-Warning "CRITICAL: Could not discover all sites: $($_.Exception.Message)."
        Write-Warning "This is likely a PERMISSIONS issue in your M365 tenant, not a script bug."
        Write-Warning "Falling back to searching the root site only."
    }

    if ($null -eq $allSites -or $allSites.Count -eq 0) {
        try {
            # --- CORRECTED CODE ---
            # Also removed the invalid -Property parameter from the fallback call.
            $rootSite = Get-MgSite -SiteId "root" -ErrorAction Stop
            if ($rootSite) { $allSites = @($rootSite) }
        }
        catch {
             Write-Error "Failed to retrieve any sites, including the root site: $($_.Exception.Message)"
             return
        }
    }
    
    # Filter for SharePoint sites (excluding personal OneDrive sites)
    $sites = $allSites | Where-Object { $_.WebUrl -and $_.WebUrl -notlike "*/personal/*" -and $_.Id }
    Write-Host "Found $($sites.Count) SharePoint site(s) to scan." -ForegroundColor Green
    
    if ($sites.Count -eq 0) {
        Write-Warning "No SharePoint sites found to scan. Please investigate permissions."
        return
    }
    
    Write-Host "`nNow searching for video files..."
    
    $allVideosData = @()
    $videoExtensions = @('.mp4', '.mov', '.avi', '.wmv', '.mkv', '.webm', '.m4v', '.mpg', '.mpeg')
    $scanStartTime = Get-Date
    $sitesWithErrors = 0

    foreach ($site in $sites) {
        # --- CORRECTED CODE ---
        # Access the DataLocationCode from the nested SiteCollection property
        $geoLocation = if ($site.SiteCollection) { $site.SiteCollection.DataLocationCode } else { "N/A" }
        Write-Host "`nProcessing site: $($site.DisplayName) (Geo: $geoLocation)" -ForegroundColor Cyan
        
        try {
            $drives = Get-MgSiteDrive -SiteId $site.Id -All -ErrorAction Stop
            Write-Host "  Found $($drives.Count) document libraries." -ForegroundColor Gray

            foreach ($drive in $drives) {
                try {
                    Write-Host "    Scanning library: $($drive.Name)..." -NoNewline
                    
                    $allItems = Get-MgDriveItem -DriveId $drive.Id -All -Filter "file ne null" -Property "id,name,webUrl,createdBy,createdDateTime,lastModifiedDateTime,size,video" -ErrorAction Stop

                    # Filter locally for video files
                    $driveItems = $allItems | Where-Object { 
                        $_.Name -and 
                        ($videoExtensions.Contains([System.IO.Path]::GetExtension($_.Name).ToLowerInvariant()))
                    }

                    if ($driveItems) {
                        $videoCount = ($driveItems | Measure-Object).Count
                        Write-Host " Found $videoCount video(s)" -ForegroundColor Green
                        
                        foreach ($item in $driveItems) {
                            $videoData = [PSCustomObject]@{
                                VideoName             = $item.Name
                                VideoURL              = $item.WebUrl
                                SiteName              = $site.DisplayName
                                SiteGeoLocation       = $geoLocation # <-- NEW MULTI-GEO PROPERTY
                                SiteURL               = $site.WebUrl
                                LibraryName           = $drive.Name
                                VideoCreatorName      = if($item.CreatedBy.User.DisplayName) { $item.CreatedBy.User.DisplayName } else { 'Unknown' }
                                VideoCreatorEmail     = if($item.CreatedBy.User.Email) { $item.CreatedBy.User.Email } else { 'Unknown' }
                                VideoCreationDate     = $item.CreatedDateTime
                                VideoModificationDate = $item.LastModifiedDateTime
                                VideoSizeMB           = if ($item.Size) { [math]::Round($item.Size / 1MB, 2) } else { 0 }
                                VideoDurationSeconds  = if ($item.Video -and $item.Video.Duration) { [math]::Round($item.Video.Duration / 1000, 0) } else { 0 }
                            }
                            $allVideosData += $videoData
                        }
                    } else {
                        Write-Host " No videos." -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host " Error: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        catch {
            Write-Warning "An error occurred processing site '$($site.DisplayName)': $($_.Exception.Message)"
            $sitesWithErrors++
        }
    }
    
    # Scan Summary
    $totalScanTime = [math]::Round(((Get-Date) - $scanStartTime).TotalMinutes, 1)
    Write-Host "`n" + ("-"*50)
    Write-Host "Scan Summary:" -ForegroundColor Yellow
    Write-Host "  - Total scan time: $totalScanTime minutes"
    Write-Host "  - Sites with errors: $sitesWithErrors"
    Write-Host "  - Total videos found: $($allVideosData.Count)" -ForegroundColor Green

    # 5. Export to CSV
    Write-Host ("-"*50)
    Write-Host "Step 5: Generating Report..." -ForegroundColor Yellow
    if ($allVideosData.Count -gt 0) {
        $datestring = (Get-Date).ToString("yyyyMMdd-HHmm")
        $csvfileName = Join-Path -Path $reportFolder -ChildPath "M365_VideoReport_$datestring.csv"
        
        $allVideosData | Export-Csv -Path $csvfileName -NoTypeInformation -Encoding UTF8
        Write-Host "`nReport successfully generated!" -ForegroundColor Green
        Write-Host "Report saved to: $csvfileName"
    } else {
        Write-Warning "No video files were found."
    }
}
catch {
    # 6. Global Error Catching
    Write-Error "A critical, unrecoverable error occurred: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.ScriptStackTrace)"
}
finally {
    # 7. Disconnect
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Write-Host "`nDisconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected successfully." -ForegroundColor Green
    }
    Write-Host "`nScript execution completed." -ForegroundColor Cyan
}
