<#
.SYNOPSIS
    Diagnoses and fixes Microsoft Graph PowerShell module version conflicts
.DESCRIPTION
    This script will identify version mismatches in Microsoft Graph modules and help resolve them
#>

Write-Host "Microsoft Graph Module Diagnostic and Repair Tool" -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

# 1. Check all installed Microsoft Graph modules
Write-Host "`nStep 1: Checking installed Microsoft Graph modules..." -ForegroundColor Yellow
$graphModules = Get-Module -Name "Microsoft.Graph*" -ListAvailable | 
    Select-Object Name, Version, Path | 
    Sort-Object Name, Version -Descending

if ($graphModules) {
    Write-Host "Found the following Microsoft Graph modules:" -ForegroundColor Green
    $graphModules | Format-Table -AutoSize
    
    # Check for version conflicts
    $coreModule = $graphModules | Where-Object { $_.Name -eq "Microsoft.Graph.Core" }
    $authModule = $graphModules | Where-Object { $_.Name -eq "Microsoft.Graph.Authentication" }
    
    if ($coreModule -and $authModule) {
        $coreVersions = @($coreModule.Version | Select-Object -Unique)
        $authVersions = @($authModule.Version | Select-Object -Unique)
        
        Write-Host "`nCore Module Versions: $($coreVersions -join ', ')" -ForegroundColor Yellow
        Write-Host "Auth Module Versions: $($authVersions -join ', ')" -ForegroundColor Yellow
        
        if ($coreVersions.Count -gt 1 -or $authVersions.Count -gt 1) {
            Write-Warning "Multiple versions detected! This is likely causing the conflict."
        }
    }
} else {
    Write-Warning "No Microsoft Graph modules found installed."
}

# 2. Check loaded modules
Write-Host "`nStep 2: Checking currently loaded modules..." -ForegroundColor Yellow
$loadedModules = Get-Module -Name "Microsoft.Graph*" | Select-Object Name, Version
if ($loadedModules) {
    Write-Host "Currently loaded modules:" -ForegroundColor Green
    $loadedModules | Format-Table -AutoSize
} else {
    Write-Host "No Microsoft Graph modules currently loaded." -ForegroundColor Gray
}

# 3. Offer to fix
Write-Host "`nStep 3: Resolution Options" -ForegroundColor Yellow
Write-Host "Choose an option to proceed:" -ForegroundColor White
Write-Host "  1. Remove ALL Microsoft Graph modules and reinstall fresh (Recommended)" -ForegroundColor Green
Write-Host "  2. Try to update existing modules to latest version" -ForegroundColor Yellow
Write-Host "  3. Exit without making changes" -ForegroundColor Gray

$choice = Read-Host "`nEnter your choice (1-3)"

switch ($choice) {
    "1" {
        Write-Host "`nRemoving all Microsoft Graph modules..." -ForegroundColor Yellow
        
        # First, remove from current session
        Get-Module -Name "Microsoft.Graph*" | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # Get all installed versions
        $allGraphModules = Get-Module -Name "Microsoft.Graph*" -ListAvailable
        
        foreach ($module in $allGraphModules) {
            Write-Host "  Removing $($module.Name) v$($module.Version)..." -NoNewline
            try {
                Uninstall-Module -Name $module.Name -RequiredVersion $module.Version -Force -ErrorAction Stop
                Write-Host " Done" -ForegroundColor Green
            }
            catch {
                Write-Host " Failed (may need admin rights)" -ForegroundColor Red
            }
        }
        
        Write-Host "`nInstalling fresh Microsoft Graph module suite..." -ForegroundColor Yellow
        Write-Host "This may take 3-5 minutes. Please be patient..." -ForegroundColor Cyan
        
        try {
            # Install the main module which includes all dependencies
            Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
            Write-Host "`nInstallation completed successfully!" -ForegroundColor Green
            
            # Import the modules
            Write-Host "`nImporting modules..." -ForegroundColor Yellow
            Import-Module Microsoft.Graph.Authentication -Force
            Import-Module Microsoft.Graph.Sites -Force
            Import-Module Microsoft.Graph.Files -Force
            
            Write-Host "Modules imported successfully!" -ForegroundColor Green
            
            # Test the connection
            Write-Host "`nTesting connection..." -ForegroundColor Yellow
            Connect-MgGraph -Scopes "User.Read" -NoWelcome
            $me = Get-MgUser -UserId "me" -ErrorAction Stop
            Write-Host "Success! Connected as: $($me.UserPrincipalName)" -ForegroundColor Green
            Disconnect-MgGraph
            
        }
        catch {
            Write-Error "Installation or test failed: $_"
        }
    }
    
    "2" {
        Write-Host "`nUpdating Microsoft Graph modules..." -ForegroundColor Yellow
        try {
            Update-Module -Name Microsoft.Graph -Force
            Write-Host "Update completed!" -ForegroundColor Green
            Write-Host "Please restart PowerShell and try again." -ForegroundColor Yellow
        }
        catch {
            Write-Error "Update failed: $_"
        }
    }
    
    "3" {
        Write-Host "Exiting without changes." -ForegroundColor Gray
    }
    
    default {
        Write-Host "Invalid choice. Exiting." -ForegroundColor Red
    }
}

Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
Write-Host "Diagnostic complete. If option 1 was successful, you can now run your Stream discovery script." -ForegroundColor Cyan
Write-Host "If issues persist, try running PowerShell as Administrator and choosing option 1 again." -ForegroundColor Yellow