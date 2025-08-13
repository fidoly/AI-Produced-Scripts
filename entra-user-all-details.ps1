#Requires -Version 5.1
<#
.SYNOPSIS
    Retrieves Entra ID (Azure AD) user details based on email domain with last interactive sign-in time.

.DESCRIPTION
    This script connects to Microsoft Graph API to retrieve all users from a specified domain,
    including their profile information and last interactive sign-in timestamp.
    Results are exported to a CSV file.

.PARAMETER None
    Script will prompt for all required information interactively.

.EXAMPLE
    .\Get-EntraUserDetails.ps1

.NOTES
    Version:        1.0.0
    Author:         System Administrator
    Creation Date:  2025
    Purpose/Change: Initial script development
#>

# Script Version
$ScriptVersion = "1.0.0"

# Clear console and display header
Clear-Host
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║          Entra ID User Details Export Tool v$ScriptVersion          ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Function to install required modules
function Install-RequiredModule {
    param (
        [string]$ModuleName
    )
    
    Write-Host "🔍 Checking for $ModuleName module..." -ForegroundColor Yellow
    
    if (!(Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "📦 Installing $ModuleName module..." -ForegroundColor Yellow
        Write-Host "⏳ This may take a few minutes on first installation..." -ForegroundColor Gray
        
        try {
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
            Write-Host "✅ $ModuleName module installed successfully!" -ForegroundColor Green
        }
        catch {
            Write-Host "❌ Failed to install $ModuleName module: $_" -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "✅ $ModuleName module is already installed." -ForegroundColor Green
    }
    
    return $true
}

# Function to show progress
function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Activity
    )
    
    # Ensure we don't exceed 100%
    $percentComplete = [Math]::Min([Math]::Round(($Current / $Total) * 100, 2), 100)
    
    # Only show progress if we have a valid percentage
    if ($percentComplete -ge 0 -and $percentComplete -le 100) {
        Write-Progress -Activity $Activity -Status "$Current of $Total users processed" -PercentComplete $percentComplete
    }
}

# Main script execution
try {
    # Install and import required modules
    Write-Host "📋 Preparing environment..." -ForegroundColor Cyan
    Write-Host ""
    
    if (!(Install-RequiredModule -ModuleName "Microsoft.Graph")) {
        throw "Failed to install Microsoft Graph module"
    }
    
    # Import module
    Write-Host "📥 Importing Microsoft Graph module..." -ForegroundColor Yellow
    Write-Host "⏳ This may take a moment, especially on first load..." -ForegroundColor Gray
    
    $importStart = Get-Date
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Reports -ErrorAction Stop
    $importEnd = Get-Date
    
    $importTime = [Math]::Round(($importEnd - $importStart).TotalSeconds, 2)
    Write-Host "✅ Modules imported successfully in $importTime seconds!" -ForegroundColor Green
    Write-Host ""
    
    # Get user domain
    Write-Host "🌐 Domain Configuration" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    do {
        $domain = Read-Host "Enter the email domain to search (e.g., contoso.com)"
        if ([string]::IsNullOrWhiteSpace($domain)) {
            Write-Host "❌ Domain cannot be empty. Please try again." -ForegroundColor Red
        }
    } while ([string]::IsNullOrWhiteSpace($domain))
    
    Write-Host ""
    
    # Get output location
    Write-Host "💾 Output Configuration" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    $defaultPath = "C:\temp"
    $outputPath = Read-Host "Enter the folder path for CSV output (default: $defaultPath)"
    
    if ([string]::IsNullOrWhiteSpace($outputPath)) {
        $outputPath = $defaultPath
    }
    
    # Create directory if it doesn't exist
    if (!(Test-Path -Path $outputPath)) {
        Write-Host "📁 Creating directory: $outputPath" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
    }
    
    # Generate filename
    $dateString = Get-Date -Format "yyyy-MM-dd"
    $fileName = "User Details for $domain - $dateString.csv"
    $fullPath = Join-Path -Path $outputPath -ChildPath $fileName
    
    Write-Host "📄 Output file: $fullPath" -ForegroundColor Green
    Write-Host ""
    
    # Connect to Microsoft Graph
    Write-Host "🔐 Authentication" -ForegroundColor Cyan
    Write-Host "━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
    Write-Host "🌐 Launching browser for authentication..." -ForegroundColor Yellow
    Write-Host "⚠️  Please sign in with an account that has User.Read.All permissions" -ForegroundColor Yellow
    
    try {
        Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All", "Reports.Read.All", "Directory.Read.All" -NoWelcome
        Write-Host "✅ Successfully authenticated!" -ForegroundColor Green
    }
    catch {
        throw "Authentication failed: $_"
    }
    
    Write-Host ""
    Write-Host "🔍 Searching for users with domain: $domain" -ForegroundColor Cyan
    Write-Host ""
    
    # Build filter query
    $filter = "endsWith(mail,'@$domain') or endsWith(userPrincipalName,'@$domain')"
    
    # Get total user count first
    Write-Host "📊 Counting users..." -ForegroundColor Yellow
    
    # Get the count using a different approach to avoid variable conflicts
    try {
        # First attempt: Try to get count using ConsistencyLevel header
        $countRequest = Get-MgUser -Filter $filter -ConsistencyLevel eventual -Top 1 -CountVariable UserCount
        $userCount = $UserCount
        
        # If CountVariable didn't work, get all users and count them
        if (-not $userCount -or $userCount -eq 0) {
            Write-Host "   ⏳ Getting user count (this may take a moment)..." -ForegroundColor Gray
            $allUsers = Get-MgUser -Filter $filter -All -ConsistencyLevel eventual -Property Id
            $userCount = $allUsers.Count
        }
    }
    catch {
        Write-Host "   ⚠️  Count query failed, retrieving all users to count..." -ForegroundColor Yellow
        $allUsers = Get-MgUser -Filter $filter -All -ConsistencyLevel eventual -Property Id
        $userCount = $allUsers.Count
    }
    
    if ($userCount -eq 0) {
        Write-Host "⚠️  No users found with domain: $domain" -ForegroundColor Yellow
        return
    }
    
    Write-Host "✅ Found $userCount users to process" -ForegroundColor Green
    Write-Host ""
    
    # Retrieve users with all properties using -All parameter to handle pagination
    Write-Host "📥 Retrieving user data (this may take a moment for large datasets)..." -ForegroundColor Cyan
    
    # Use -All to ensure we get all users, not just first page
    $users = @()
    $pageSize = 999  # Maximum page size for Graph API
    
    try {
        $users = Get-MgUser -Filter $filter -All -Property * -ConsistencyLevel eventual -PageSize $pageSize
        Write-Host "✅ Successfully retrieved $($users.Count) users" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Failed to retrieve all users at once, using pagination..." -ForegroundColor Yellow
        
        # Fallback to manual pagination if needed
        $users = @()
        $nextLink = $null
        $pageNumber = 1
        
        do {
            Write-Host "   📄 Retrieving page $pageNumber..." -ForegroundColor Gray
            
            if ($nextLink) {
                $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
            }
            else {
                $uri = "https://graph.microsoft.com/v1.0/users?`$filter=$filter&`$select=*&`$top=$pageSize&`$count=true"
                $response = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{ 'ConsistencyLevel' = 'eventual' }
            }
            
            $users += $response.value
            $nextLink = $response.'@odata.nextLink'
            $pageNumber++
            
        } while ($nextLink)
        
        Write-Host "✅ Successfully retrieved $($users.Count) users across $($pageNumber - 1) pages" -ForegroundColor Green
    }
    
    # Process users and get sign-in data
    Write-Host "🔄 Processing user details and sign-in information..." -ForegroundColor Cyan
    $processedUsers = @()
    $currentUser = 0
    $totalUsers = $users.Count
    
    foreach ($user in $users) {
        $currentUser++
        Show-Progress -Current $currentUser -Total $totalUsers -Activity "Processing users"
        
        try {
            # Get sign-in activity using the beta endpoint for more reliable data
            $signInActivity = $null
            try {
                # Use the beta endpoint which has better sign-in data
                $betaUri = "https://graph.microsoft.com/beta/users/$($user.Id)?`$select=signInActivity"
                $signInResponse = Invoke-MgGraphRequest -Uri $betaUri -Method GET
                $signInActivity = $signInResponse.signInActivity
            }
            catch {
                # Fallback to user object's sign-in activity
                $signInActivity = $user.SignInActivity
            }
            
            # Get manager information
            $managerName = ""
            if ($user.Manager -and $user.Manager.Id) {
                try {
                    $manager = Get-MgUser -UserId $user.Manager.Id -Property DisplayName -ErrorAction SilentlyContinue
                    $managerName = $manager.DisplayName
                }
                catch {
                    $managerName = ""
                }
            }
            
            # Get license names
            $licenseNames = @()
            if ($user.AssignedLicenses) {
                foreach ($license in $user.AssignedLicenses) {
                    # Common license SKU translations
                    $skuName = switch ($license.SkuId) {
                        "06ebc4ee-1bb5-47dd-8120-11324bc54e06" { "E5" }
                        "c7df2760-2c81-4ef7-b578-5b5392b571df" { "E5" }
                        "26d45bd9-adf1-46cd-a9e1-51e9a5524128" { "E3" }
                        "189a915c-fe4f-4ffa-bde4-85b9628d07a0" { "Developer E5" }
                        "b05e124f-c7cc-45a0-a6aa-8cf78c946968" { "F1" }
                        "4b585984-651b-448a-9e53-3b10f069cf7f" { "F3" }
                        "18181a46-0d4e-45cd-891e-60aabd171b4e" { "Office 365 E1" }
                        "6fd2c87f-b296-42f0-b197-1e91e994b900" { "Office 365 E3" }
                        "c7df2760-2c81-4ef7-b578-5b5392b571df" { "Office 365 E5" }
                        "29a2f828-8f39-4837-b8ff-c957e86abe3c" { "Business Basic" }
                        "f245ecc8-75af-4f8e-b61f-27d8114de5f3" { "Business Standard" }
                        "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" { "Business Premium" }
                        default { $license.SkuId }
                    }
                    $licenseNames += $skuName
                }
            }
            
            # Create custom object with all user properties
            $userDetails = [PSCustomObject]@{
                'Display Name'             = $user.DisplayName
                'User Principal Name'      = $user.UserPrincipalName
                'Email'                    = $user.Mail
                'Account Enabled'          = if ($null -ne $user.AccountEnabled) { 
                    if ($user.AccountEnabled) { "Enabled" } else { "Disabled" }
                } else { "Unknown" }
                'Last Interactive Sign-In' = if ($signInActivity -and $signInActivity.lastSignInDateTime) { 
                    [DateTime]::Parse($signInActivity.lastSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                } else { "Never" }
                'Last Non-Interactive Sign-In' = if ($signInActivity -and $signInActivity.lastNonInteractiveSignInDateTime) { 
                    [DateTime]::Parse($signInActivity.lastNonInteractiveSignInDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                } else { "Never" }
                'Job Title'               = $user.JobTitle
                'Department'              = $user.Department
                'Office Location'         = $user.OfficeLocation
                'Manager'                 = $managerName
                'Mobile Phone'            = $user.MobilePhone
                'Business Phone'          = $user.BusinessPhones -join '; '
                'Street Address'          = $user.StreetAddress
                'City'                    = $user.City
                'State'                   = $user.State
                'Country'                 = $user.Country
                'Postal Code'             = $user.PostalCode
                'Employee ID'             = $user.EmployeeId
                'Employee Type'           = $user.EmployeeType
                'Created DateTime'        = if ($user.CreatedDateTime) { 
                    [DateTime]::Parse($user.CreatedDateTime).ToString("yyyy-MM-dd HH:mm:ss")
                } else { "Unknown" }
                'License Details'         = $licenseNames -join '; '
                'Proxy Addresses'         = ($user.ProxyAddresses -join '; ')
                'Account Type'            = $user.UserType
                'Usage Location'          = $user.UsageLocation
                'Company Name'            = $user.CompanyName
                'Object ID'               = $user.Id
            }
            
            $processedUsers += $userDetails
        }
        catch {
            Write-Host "`n⚠️  Warning: Could not process user $($user.UserPrincipalName): $_" -ForegroundColor Yellow
        }
    }
    
    Write-Progress -Activity "Processing users" -Completed
    Write-Host ""
    
    # Export to CSV
    Write-Host "💾 Exporting data to CSV..." -ForegroundColor Cyan
    try {
        $processedUsers | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8
        Write-Host "✅ Successfully exported $($processedUsers.Count) users to:" -ForegroundColor Green
        Write-Host "   📄 $fullPath" -ForegroundColor White
    }
    catch {
        throw "Failed to export CSV: $_"
    }
    
    # Summary
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                    ✅ Export Complete!                         ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "📊 Summary:" -ForegroundColor Cyan
    Write-Host "   • Domain: $domain" -ForegroundColor White
    Write-Host "   • Total Users Processed: $($processedUsers.Count)" -ForegroundColor White
    Write-Host "   • Output File: $fileName" -ForegroundColor White
    Write-Host "   • Location: $outputPath" -ForegroundColor White
    Write-Host ""
    
    # Open file location
    $openFolder = Read-Host "Would you like to open the output folder? (Y/N)"
    if ($openFolder -eq 'Y' -or $openFolder -eq 'y') {
        Start-Process explorer.exe -ArgumentList $outputPath
    }
}
catch {
    Write-Host ""
    Write-Host "❌ Script Error: $_" -ForegroundColor Red
    Write-Host "📋 Stack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
}
finally {
    # Disconnect from Microsoft Graph
    if (Get-MgContext) {
        Write-Host ""
        Write-Host "🔒 Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
        Disconnect-MgGraph | Out-Null
        Write-Host "✅ Disconnected successfully." -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}