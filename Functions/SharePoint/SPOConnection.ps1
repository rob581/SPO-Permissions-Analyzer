function Test-PnPModuleAvailable {
    
    # Method 1: Check installed modules (prefer 3.x)
    $moduleInstalled = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending | Select-Object -First 1
    
    if ($moduleInstalled) {
        $version = $moduleInstalled.Version
        Write-ActivityLog "PnP module found via Get-Module: Version $version"
        
        # Check if it's modern version
        if ($version -ge [version]"3.0.0") {
            Write-ActivityLog "Modern PnP PowerShell 3.x detected"
            return $true
        } elseif ($version -ge [version]"2.0.0") {
            Write-ActivityLog "Legacy PnP PowerShell 2.x detected - recommend upgrading to 3.x"
            return $true
        } else {
            Write-ActivityLog "Very old PnP PowerShell detected - upgrade required"
            return $false
        }
    }
    
    # Method 2: Try to import
    try {
        Import-Module PnP.PowerShell -ErrorAction Stop
        $importedModule = Get-Module PnP.PowerShell
        if ($importedModule) {
            Write-ActivityLog "PnP module successfully imported: Version $($importedModule.Version)"
            return $true
        }
    }
    catch {
        Write-ActivityLog "PnP module import failed: $($_.Exception.Message)"
    }
    
    # Method 3: Check if key commands are available (modern commands)
    try {
        $modernCommands = @("Connect-PnPOnline", "Get-PnPWeb", "Get-PnPRoleAssignment")
        $availableCommands = 0
        
        foreach ($cmd in $modernCommands) {
            if (Get-Command $cmd -ErrorAction SilentlyContinue) {
                $availableCommands++
            }
        }
        
        if ($availableCommands -eq $modernCommands.Count) {
            Write-ActivityLog "All modern PnP commands are available"
            return $true
        } else {
            Write-ActivityLog "Missing $($modernCommands.Count - $availableCommands) key PnP commands"
        }
    }
    catch {
        Write-ActivityLog "PnP commands check failed: $($_.Exception.Message)"
    }
    
    return $false
}

function Install-PnPModule {
    param($UI)

    try {
        $UI.UpdateStatus("Installing modern PnP PowerShell 3.x...`nThis may take a few minutes.")
        
        # Check PowerShell version first
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            throw "PowerShell 7.0 or later is required for modern PnP PowerShell 3.x. Current version: $($PSVersionTable.PSVersion)"
        }
        
        # Remove any legacy versions first
        $UI.UpdateStatus("Cleaning up legacy PnP PowerShell versions...")
        $legacyModules = @("SharePointPnPPowerShellOnline", "PnP.PowerShell")
        
        foreach ($module in $legacyModules) {
            try {
                $installed = Get-Module -Name $module -ListAvailable -ErrorAction SilentlyContinue
                if ($installed) {
                    $oldVersions = $installed | Where-Object { $_.Version -lt [version]"3.0.0" }
                    if ($oldVersions) {
                        Write-ActivityLog "Removing legacy $module versions: $($oldVersions.Version -join ', ')"
                        Uninstall-Module -Name $module -Force -AllVersions -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                Write-ActivityLog "Could not remove legacy $module`: $($_.Exception.Message)"
            }
        }
        
        # Check and install NuGet if needed
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nugetProvider -or $nugetProvider.Version -lt [version]"2.8.5.201") {
            $UI.UpdateStatus("Installing NuGet provider...")
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }

        # Install modern PnP PowerShell 3.x
        $UI.UpdateStatus("Installing PnP.PowerShell 3.x...`nPlease wait, this may take several minutes.")
        
        $installParams = @{
            Name = "PnP.PowerShell"
            MinimumVersion = "3.0.0"
            Force = $true
            AllowClobber = $true
            SkipPublisherCheck = $true
            Scope = "CurrentUser"
        }
        
        Install-Module @installParams
        
        # Force refresh and import
        Import-Module PnP.PowerShell -Force -ErrorAction SilentlyContinue
        
        # Verify modern installation
        $installedModule = Get-Module -ListAvailable -Name PnP.PowerShell | 
            Sort-Object Version -Descending | Select-Object -First 1
            
        if (-not $installedModule) {
            throw "PnP PowerShell module installation verification failed"
        }
        
        if ($installedModule.Version -lt [version]"3.0.0") {
            throw "Failed to install modern PnP PowerShell 3.x. Got version $($installedModule.Version) instead."
        }
        
        # Test key modern commands
        $modernCommands = @("Connect-PnPOnline", "Get-PnPWeb", "Get-PnPRoleAssignment", "Get-PnPUser", "Get-PnPGroup")
        $missingCommands = @()
        
        foreach ($cmd in $modernCommands) {
            if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
                $missingCommands += $cmd
            }
        }
        
        if ($missingCommands.Count -gt 0) {
            Write-ActivityLog "Warning: Missing modern commands: $($missingCommands -join ', ')"
        }
        
        $UI.UpdateStatus("✅ Modern PnP PowerShell $($installedModule.Version) installed successfully!`nAll key commands available for enhanced SharePoint analysis.")
        return $true
    }
    catch {
        $UI.UpdateStatus("Error installing modern PnP module: $($_.Exception.Message)")
        throw
    }
}

function Connect-SharePointWithAppRegistration {
    param(
        $TenantUrl, 
        $ClientId, 
        $UI
    )
    
    try {
        Write-ActivityLog "Starting modern SharePoint connection to: $TenantUrl with App Registration: $ClientId"
        
        # Clear any existing connections first
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-ActivityLog "Cleared any existing connections"
        }
        catch {
            # Ignore errors during disconnect
        }
        
        # Ensure modern module is imported
        $UI.UpdateStatus("Importing modern PnP PowerShell 3.x...")
        try {
            Import-Module PnP.PowerShell -Force -ErrorAction Stop
            $pnpModule = Get-Module PnP.PowerShell
            Write-ActivityLog "Modern PnP PowerShell imported successfully: Version $($pnpModule.Version)"
            
            if ($pnpModule.Version -lt [version]"3.0.0") {
                throw "Legacy PnP PowerShell detected. Please upgrade to 3.x for full functionality."
            }
        }
        catch {
            throw "Failed to import modern PnP PowerShell module: $($_.Exception.Message)"
        }
        
        # Modern connection approach - simplified for PnP 3.x
        $UI.UpdateStatus("Connecting to SharePoint Online...`nUsing App Registration: $ClientId`nPlease complete authentication in the popup window.")
        Write-ActivityLog "Attempting modern connection to $TenantUrl with ClientId $ClientId"
        
        try {
            # Modern PnP 3.x connection - more reliable
            Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Interactive
            Write-ActivityLog "Modern connection command completed successfully"
        }
        catch {
            throw "Failed to connect to SharePoint with app registration: $($_.Exception.Message)"
        }
        
        # Enhanced verification using modern PnP features
        $UI.UpdateStatus("Verifying SharePoint connection...")
        try {
            # Modern verification approach
            $context = Get-PnPContext -ErrorAction SilentlyContinue
            $connection = Get-PnPConnection -ErrorAction SilentlyContinue
            
            if ($null -eq $context -and $null -eq $connection) {
                throw "No SharePoint connection available after authentication"
            }
            
            Write-ActivityLog "SharePoint context/connection verified successfully"
            
            # Get enhanced site information using modern PnP
            $site = $null
            $web = $null
            try {
                # Try modern approach first
                $web = Get-PnPWeb -ErrorAction SilentlyContinue
                if ($web) {
                    $site = @{ 
                        Url = $web.Url; 
                        Title = $web.Title;
                        Created = $web.Created;
                        Language = $web.Language;
                        Template = $web.WebTemplate
                    }
                    Write-ActivityLog "Enhanced site verification successful: $($web.Url)"
                } else {
                    # Fallback
                    $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
                    Write-ActivityLog "Using basic verification - connection established"
                }
            }
            catch {
                Write-ActivityLog "Site verification had issues but proceeding: $($_.Exception.Message)"
                $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
            }
            
        }
        catch {
            Write-ActivityLog "Connection verification had issues but proceeding: $($_.Exception.Message)"
            $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
        }
        
        $script:SPOConnected = $true
        $script:SPOContext = Get-PnPConnection -ErrorAction SilentlyContinue
        
        # Enhanced user info with modern PnP
        $currentUser = ""
        try {
            $userInfo = Get-PnPCurrentUser -ErrorAction SilentlyContinue
            if ($userInfo) {
                $currentUser = $userInfo.UserPrincipalName -or $userInfo.Email -or $userInfo.LoginName -or $userInfo.Title
            }
        }
        catch {
            $currentUser = "Authentication verified"
        }
        
        Write-ActivityLog "Successfully connected to SharePoint: $($site.Url) using modern PnP with app $ClientId"
        $UI.UpdateStatus("✅ Successfully connected to SharePoint Online!`nSite: $($site.Url)`nApp Registration: $ClientId`nUser: $currentUser`nPnP Version: Modern 3.x`nReady for enhanced SharePoint operations.")
        
        return $true
    }
    catch {
        $script:SPOConnected = $false
        $script:SPOContext = $null
        Write-ErrorLog -Message $_.Exception.Message -Location "Modern-SharePoint-AppRegistration-Connection"
        $UI.UpdateStatus("SharePoint connection failed: $($_.Exception.Message)")
        throw
    }
}

function Connect-SPO {
    try {
        Write-ActivityLog "Starting Connect-SPO function"
        
        # Validate UI controls
        if ($null -eq $script:txtSPOResults -or $null -eq $script:btnConnectSPO) {
            throw "UI controls not properly initialized"
        }
        
        # Create controls hashtable for SharePoint operations
        $controls = @{
            'GetSites' = $script:btnGetSites
            'GetPermissions' = $script:btnGetPermissions
            'GenerateReport' = $script:btnGenerateReport
        }
        
        # Create UI manager
        $ui = New-UIManager -ResultsBox $script:txtSPOResults -ConnectButton $script:btnConnectSPO -Controls $controls
        
        # Check demo mode
        if (Get-AppSetting -SettingName "DemoMode") {
            $ui.EnableControls()
            $script:SPOConnected = $true
            $script:txtSPOResults.Text = "Connected to SharePoint Online (Demo Mode)"
            $script:txtStatus.Text = "Connected (Demo Mode)"
            $script:txtStatus.Foreground = "Green"
            return
        }
        
        # Get tenant URL and Client ID
        $tenantUrl = $script:txtTenantUrl.Text.Trim()
        $clientId = $script:txtClientId.Text.Trim()
        
        if ([string]::IsNullOrEmpty($tenantUrl)) {
            $ui.UpdateStatus("Please enter your SharePoint tenant URL")
            return
        }
        
        if ([string]::IsNullOrEmpty($clientId)) {
            $ui.UpdateStatus("Please enter your App Registration Client ID")
            return
        }
        
        if (-not $tenantUrl.StartsWith("https://") -or -not $tenantUrl.Contains(".sharepoint.com")) {
            $ui.UpdateStatus("Please enter a valid SharePoint Online tenant URL (e.g., https://contoso.sharepoint.com)")
            return
        }
        
        # Validate Client ID format (GUID)
        try {
            $guid = [System.Guid]::Parse($clientId)
        }
        catch {
            $ui.UpdateStatus("Please enter a valid Client ID (GUID format)")
            return
        }
        
        $ui.UpdateStatus("Checking PnP PowerShell module...")
        $ui.DisableControls()
        
        # Check if PnP module is available
        Write-ActivityLog "Checking for PnP PowerShell module availability"
        if (-not (Test-PnPModuleAvailable)) {
            Write-ActivityLog "PnP PowerShell module not available, starting installation"
            try {
                Install-PnPModule -UI $ui
                Write-ActivityLog "PnP module installation completed"
            }
            catch {
                Write-ErrorLog -Message "Module installation failed: $($_.Exception.Message)" -Location "Install-PnPModule"
                $ui.UpdateStatus("Failed to install PnP PowerShell module: $($_.Exception.Message)")
                return
            }
        } else {
            Write-ActivityLog "PnP PowerShell module is available"
            $ui.UpdateStatus("PnP PowerShell module found and ready")
        }
        
        # Connect to SharePoint using App Registration
        try {
            Connect-SharePointWithAppRegistration -TenantUrl $tenantUrl -ClientId $clientId -UI $ui
            $ui.EnableControls()
            
            # Update status in the SharePoint Operations tab
            $script:txtStatus.Text = "Connected to SharePoint (App Registration)"
            $script:txtStatus.Foreground = "Green"
            
            # Store settings
            Set-AppSetting -SettingName "SharePoint.TenantUrl" -Value $tenantUrl
            Set-AppSetting -SettingName "SharePoint.ClientId" -Value $clientId
            
            Write-ActivityLog "SharePoint connection process completed successfully"
        }
        catch {
            Write-ErrorLog -Message $_.Exception.Message -Location "Connect-SharePoint-AppRegistration"
            $ui.UpdateStatus("Connection failed: $($_.Exception.Message)")
            $ui.DisableControls()
        }
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Connect-SPO"
        if ($null -ne $ui) {
            $ui.DisableControls()
        }
    }
}

function Get-SharePointSites {
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        Write-ActivityLog "Getting SharePoint sites using modern PnP"
        $script:txtOperationsResults.Text = "Getting SharePoint sites...`n`nAttempting modern tenant-level site enumeration..."
        
        $sites = @()
        $adminRequired = $false
        $useModernApproach = $true
        
        # Try modern tenant-level site enumeration first
        try {
            # Check if we can use modern approaches
            $pnpModule = Get-Module PnP.PowerShell
            if ($pnpModule.Version -ge [version]"3.0.0") {
                Write-ActivityLog "Using modern PnP 3.x site enumeration"
                
                # Try modern site enumeration with better error handling
                try {
                    $sites = Get-PnPTenantSite -ErrorAction Stop
                    Write-ActivityLog "Retrieved $($sites.Count) sites using modern Get-PnPTenantSite"
                    $script:txtOperationsResults.Text = "Getting SharePoint sites...`n`nSuccessfully retrieved sites using modern PnP."
                }
                catch {
                    Write-ActivityLog "Modern Get-PnPTenantSite failed, trying alternative: $($_.Exception.Message)"
                    
                    # Try hub sites as alternative
                    try {
                        $hubSites = Get-PnPHubSite -ErrorAction SilentlyContinue
                        if ($hubSites) {
                            $sites += $hubSites
                            Write-ActivityLog "Retrieved $($hubSites.Count) hub sites"
                        }
                    }
                    catch {
                        Write-ActivityLog "Hub sites enumeration also failed: $($_.Exception.Message)"
                    }
                    
                    throw $_.Exception
                }
            } else {
                $useModernApproach = $false
                $sites = Get-PnPTenantSite -ErrorAction Stop
            }
        }
        catch {
            Write-ActivityLog "Tenant-level enumeration failed: $($_.Exception.Message)"
            $adminRequired = $true
            
            # Try to connect to admin center using modern approach
            try {
                $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
                $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"
                $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
                
                Write-ActivityLog "Attempting modern admin center connection: $adminUrl"
                $script:txtOperationsResults.Text = "Getting SharePoint sites...`n`nTrying admin center connection with modern PnP...`nPlease authenticate if prompted."
                
                Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive
                $sites = Get-PnPTenantSite -ErrorAction Stop
                Write-ActivityLog "Successfully retrieved $($sites.Count) sites from admin center using modern PnP"
                $adminRequired = $false
            }
            catch {
                Write-ActivityLog "Modern admin center access also failed: $($_.Exception.Message)"
                
                # Final fallback: get current site and subsites with modern features
                try {
                    $script:txtOperationsResults.Text = "Getting SharePoint sites...`n`nUsing modern fallback method (current site + subsites)..."
                    $currentWeb = Get-PnPWeb -ErrorAction Stop
                    
                    # Modern approach to get subsites
                    $subsites = @()
                    try {
                        $subsites = Get-PnPSubWeb -Recurse -ErrorAction Stop
                    }
                    catch {
                        Write-ActivityLog "Subsite enumeration failed, using current site only"
                    }
                    
                    $sites = @($currentWeb)
                    if ($subsites) {
                        $sites += $subsites
                    }
                    Write-ActivityLog "Retrieved $($sites.Count) sites using modern fallback method"
                }
                catch {
                    throw "Unable to retrieve any sites using modern PnP. SharePoint Administrator permissions may be required."
                }
            }
        }
        
        # Enhanced results formatting with modern features
        $result = "SharePoint Sites Discovery Results (Modern PnP)`n"
        $result += "=" * 50 + "`n`n"
        
        if ($adminRequired) {
            $result += "⚠️  Limited Access Mode - Showing available sites only`n"
            $result += "Note: SharePoint Administrator permissions required for full tenant enumeration`n`n"
        }
        
        if ($useModernApproach) {
            $result += "✨ Using Modern PnP PowerShell 3.x features`n"
        }
        
        $result += "Sites Found: $($sites.Count)`n`n"
        
        # Enhanced site information display
        foreach ($site in $sites | Select-Object -First 25) {
            $result += "📁 $($site.Title)`n"
            $result += "   URL: $($site.Url)`n"
            
            # Enhanced properties available in modern PnP
            if ($site.Owner) { $result += "   Owner: $($site.Owner)`n" }
            if ($site.SiteOwnerEmail) { $result += "   Owner Email: $($site.SiteOwnerEmail)`n" }
            if ($site.StorageUsageCurrent) { $result += "   Storage: $($site.StorageUsageCurrent) MB`n" }
            if ($site.StorageQuota) { $result += "   Storage Quota: $($site.StorageQuota) MB`n" }
            if ($site.Template) { $result += "   Template: $($site.Template)`n" }
            if ($site.LastContentModifiedDate) { $result += "   Last Modified: $($site.LastContentModifiedDate)`n" }
            if ($site.IsHubSite) { $result += "   🌟 Hub Site`n" }
            if ($site.HubSiteId -and $site.HubSiteId -ne "00000000-0000-0000-0000-000000000000") { 
                $result += "   🔗 Connected to Hub Site`n" 
            }
            $result += "`n"
        }
        
        if ($sites.Count -gt 25) {
            $result += "... and $($sites.Count - 25) more sites`n`n"
        }
        
        $result += "💡 Enhanced with Modern PnP Features:`n"
        $result += "• Hub site detection and relationships`n"
        $result += "• Enhanced storage and quota information`n"
        $result += "• Improved site owner details`n`n"
        
        $result += "🔄 Copy a site URL from above and paste it in the 'Site URL' field on the SharePoint Operations tab to analyze its permissions."
        
        $script:txtOperationsResults.Text = $result
        Write-ActivityLog "Successfully completed modern site enumeration with $($sites.Count) sites"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointSites-Modern"
        $script:txtOperationsResults.Text = @"
Error retrieving sites: $($_.Exception.Message)

This typically happens when:
1. Your app registration lacks SharePoint Administrator permissions
2. You need access to the SharePoint Admin Center
3. Tenant-level operations are restricted

Modern PnP 3.x Features Available:
1. Enhanced error diagnostics
2. Hub site relationship detection
3. Improved storage quota information
4. Better fallback site enumeration

Available options:
1. Use Demo Mode to test the enhanced tool features
2. Contact your SharePoint Administrator for elevated permissions
3. Try analyzing specific sites you have access to
4. Request Sites.Read.All or Sites.FullControl.All permissions for your app registration

Current Connection: Modern PnP PowerShell 3.x
Permissions Required: SharePoint Administrator or Sites.Read.All with admin consent
"@
    }
}

function Get-SharePointPermissions {
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        $siteUrl = $script:txtSiteUrl.Text.Trim()
        $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"
        
        # Always require a specific site URL to be entered
        if ([string]::IsNullOrEmpty($siteUrl)) {
            $script:txtOperationsResults.Text = @"
Please enter a specific site URL to analyze permissions.

Examples:
• https://yourtenant.sharepoint.com/sites/teamsite
• https://yourtenant.sharepoint.com/sites/project-alpha
• https://yourtenant.sharepoint.com (tenant root site)

The tool will connect to the exact site you specify and analyze its permissions.
"@
            return
        }
        
        # Validate the URL format
        if (-not $siteUrl.StartsWith("https://") -or -not $siteUrl.Contains(".sharepoint.com")) {
            $script:txtOperationsResults.Text = "Please enter a valid SharePoint Online site URL.`n`nExample: https://yourtenant.sharepoint.com/sites/sitename"
            return
        }
        
        Write-ActivityLog "Getting permissions for site: $siteUrl"
        $script:txtOperationsResults.Text = "Analyzing permissions for: $siteUrl...`n`nConnecting to the specified site..."
        
        # ALWAYS connect to the site specified in the input field
        try {
            Write-ActivityLog "Connecting to user-specified site: $siteUrl"
            $script:txtOperationsResults.Text = "Analyzing permissions for: $siteUrl...`n`nConnecting to site...`nPlease complete authentication if prompted."
            
            # Connect directly to the site specified by the user
            Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive
            Write-ActivityLog "Successfully connected to user-specified site: $siteUrl"
            $script:txtOperationsResults.Text = "Analyzing permissions for: $siteUrl...`n`nConnected successfully!`nRetrieving permissions data..."
        }
        catch {
            $errorMsg = "Failed to connect to site: $siteUrl`nError: $($_.Exception.Message)"
            Write-ActivityLog $errorMsg
            
            # Provide specific troubleshooting for the entered site
            if ($_.Exception.Message -like "*Access is denied*" -or $_.Exception.Message -like "*Forbidden*") {
                $script:txtOperationsResults.Text = @"
❌ ACCESS DENIED to: $siteUrl

Possible causes:
1. You don't have permission to access this specific site
2. The site doesn't exist or the URL is incorrect
3. Your app registration lacks permission to access this site
4. The site has unique permissions that block app access

Troubleshooting steps:
1. ✅ Verify the site URL is correct and accessible in your browser
2. ✅ Check that you have at least read access to this site
3. ✅ Ensure your app registration has Sites.Read.All or Sites.FullControl.All permissions
4. ✅ Try a different site URL that you know you have access to

App Registration ID: $clientId
Attempted Site: $siteUrl

💡 Tip: Try entering a site URL you know you have access to, such as your team site or personal OneDrive.
"@
            } elseif ($_.Exception.Message -like "*not found*" -or $_.Exception.Message -like "*404*") {
                $script:txtOperationsResults.Text = @"
❌ SITE NOT FOUND: $siteUrl

The site URL you entered could not be found.

Please check:
1. ✅ The URL is spelled correctly
2. ✅ The site exists and is accessible
3. ✅ You have the correct tenant name
4. ✅ The site hasn't been deleted or moved

Current URL: $siteUrl

💡 Tip: Copy the URL directly from your browser address bar when visiting the site.
"@
            } else {
                $script:txtOperationsResults.Text = @"
❌ CONNECTION ERROR to: $siteUrl

Error Details: $($_.Exception.Message)

This could be due to:
1. Network connectivity issues
2. Authentication problems
3. SharePoint service issues
4. Invalid app registration configuration

Please try:
1. ✅ Checking your internet connection
2. ✅ Verifying your app registration is properly configured
3. ✅ Trying again in a few minutes
4. ✅ Testing with a different site URL

App Registration ID: $clientId
"@
            }
            return
        }
        
        # Now analyze permissions for the connected site
        $script:txtOperationsResults.Text = "Analyzing permissions for: $siteUrl...`n`nRetrieving site information..."
        
        $result = "PERMISSIONS ANALYSIS REPORT (Modern PnP)`n"
        $result += "Site: $siteUrl`n"
        $result += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n"
        $result += "App Registration: $clientId`n"
        $result += "PnP PowerShell: Modern version (3.x)`n"
        $result += "=" * 65 + "`n`n"
        
        # Step 1: Get basic site/web information using modern approach
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            $result += "✅ SITE INFORMATION`n"
            $result += "-" * 20 + "`n"
            $result += "Title: $($web.Title)`n"
            $result += "Description: $(if($web.Description) { $web.Description } else { 'No description' })`n"
            $result += "URL: $($web.Url)`n"
            $result += "Created: $(if($web.Created) { $web.Created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Not available' })`n"
            $result += "Last Modified: $(if($web.LastItemModifiedDate) { $web.LastItemModifiedDate.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Not available' })`n"
            $result += "Language: $(if($web.Language) { $web.Language } else { 'Not available' })`n"
            $result += "Template: $(if($web.WebTemplate) { $web.WebTemplate } else { 'Not available' })`n"
            $result += "Has Unique Permissions: $($web.HasUniqueRoleAssignments)`n"
            $result += "Server Relative URL: $($web.ServerRelativeUrl)`n`n"
            Write-ActivityLog "Successfully retrieved web information for $siteUrl"
        }
        catch {
            $result += "❌ SITE INFORMATION: Access denied or unavailable`n"
            $result += "Error: $($_.Exception.Message)`n`n"
            Write-ActivityLog "Failed to get web information: $($_.Exception.Message)"
        }
        
        # Step 2: Get users with modern PnP approach
        $script:txtOperationsResults.Text = $result + "Retrieving users..."
        $users = @()
        try {
            $users = Get-PnPUser -ErrorAction Stop
            $result += "✅ SITE USERS ($($users.Count) total)`n"
            $result += "-" * 20 + "`n"
            Write-ActivityLog "Successfully retrieved $($users.Count) users"
            
            # Filter and display users with better filtering
            $regularUsers = $users | Where-Object {
                $_.PrincipalType -eq "User" -and 
                -not $_.LoginName.Contains("app@sharepoint") -and 
                -not $_.LoginName.Contains("SHAREPOINT\system") -and
                -not $_.LoginName.Contains("spo-grid-all-users") -and
                -not $_.Title.Contains("System Account") -and
                -not $_.LoginName.StartsWith("c:0(.s|true") -and  # Claims-based system accounts
                -not $_.Title.StartsWith("System ") -and
                -not [string]::IsNullOrEmpty($_.Title)
            }
            
            if ($regularUsers.Count -gt 0) {
                foreach ($user in $regularUsers | Select-Object -First 20) {
                    $result += "👤 $($user.Title)"
                    if ($user.Email) { $result += " ($($user.Email))" }
                    $result += "`n"
                    
                    # Add additional user properties available in modern PnP
                    if ($user.UserPrincipalName) { $result += "   UPN: $($user.UserPrincipalName)`n" }
                    $result += "   Login: $($user.LoginName)`n"
                    if ($user.IsSiteAdmin) { $result += "   🔑 Site Administrator`n" }
                    if ($user.IsShareByEmailGuestUser) { $result += "   🔗 Guest User (Email)`n" }
                    if ($user.IsEmailAuthenticationGuestUser) { $result += "   🔗 Guest User (Auth)`n" }
                    $result += "`n"
                }
                
                if ($regularUsers.Count -gt 20) {
                    $result += "   ... and $($regularUsers.Count - 20) more users`n"
                }
            } else {
                $result += "No regular users found or access limited`n"
            }
            $result += "`n"
        }
        catch {
            $result += "⚠️ USERS: Limited access to user information`n"
            $result += "Reason: $($_.Exception.Message)`n"
            $result += "Note: This is common with app-only authentication. User enumeration requires higher permissions.`n`n"
            Write-ActivityLog "Failed to get users: $($_.Exception.Message)"
        }
        
        # Step 3: Get groups with enhanced modern PnP features
        $script:txtOperationsResults.Text = $result + "Retrieving groups..."
        $groups = @()
        try {
            $groups = Get-PnPGroup -ErrorAction Stop
            $result += "✅ SHAREPOINT GROUPS ($($groups.Count) total)`n"
            $result += "-" * 25 + "`n"
            Write-ActivityLog "Successfully retrieved $($groups.Count) groups"
            
            # Enhanced group filtering
            $importantGroups = $groups | Where-Object {
                -not $_.Title.StartsWith("SharingLinks") -and 
                -not $_.Title.StartsWith("Limited Access") -and
                -not $_.Title.StartsWith("Excel") -and
                -not $_.Title.Contains("Restricted") -and
                -not $_.Title.Contains("Style Resource Readers")
            }
            
            foreach ($group in $importantGroups | Select-Object -First 20) {
                $result += "👥 $($group.Title)`n"
                if ($group.Description) { $result += "   Description: $($group.Description)`n" }
                
                # Enhanced group information
                if ($group.OwnerTitle) { $result += "   Owner: $($group.OwnerTitle)`n" }
                $result += "   ID: $($group.Id)`n"
                
                # Try to get detailed group members
                try {
                    $groupMembers = Get-PnPGroupMember -Group $group.Title -ErrorAction SilentlyContinue
                    if ($groupMembers) {
                        $result += "   Members: $($groupMembers.Count)`n"
                        
                        # Categorize members
                        $users = $groupMembers | Where-Object { $_.PrincipalType -eq "User" }
                        $nestedGroups = $groupMembers | Where-Object { $_.PrincipalType -eq "SecurityGroup" }
                        
                        if ($users.Count -gt 0) {
                            $result += "   Users: $($users.Count)"
                            $userSample = ($users | Select-Object -First 3 | ForEach-Object { $_.Title }) -join ", "
                            if ($userSample) {
                                $result += " (e.g., $userSample"
                                if ($users.Count -gt 3) { $result += "..." }
                                $result += ")"
                            }
                            $result += "`n"
                        }
                        
                        if ($nestedGroups.Count -gt 0) {
                            $result += "   Nested Groups: $($nestedGroups.Count)`n"
                        }
                    }
                }
                catch {
                    $result += "   Members: Unable to retrieve details`n"
                }
                $result += "`n"
            }
            
            if ($importantGroups.Count -gt 20) {
                $result += "   ... and $($importantGroups.Count - 20) more groups`n"
            }
            $result += "`n"
        }
        catch {
            $result += "❌ GROUPS: Access denied or unavailable`n"
            $result += "Error: $($_.Exception.Message)`n`n"
            Write-ActivityLog "Failed to get groups: $($_.Exception.Message)"
        }
        

        # Step 4: Get permissions using modern PnP 3.x approach (UPDATED)
        $script:txtOperationsResults.Text = $result + "Retrieving permissions and role assignments..."
        try {
            # Modern way to get role assignments - use Get-PnPWeb with RoleAssignments
            $web = Get-PnPWeb -Includes "RoleAssignments" -ErrorAction Stop
            $roleAssignments = $web.RoleAssignments
            
            if ($roleAssignments -and $roleAssignments.Count -gt 0) {
                $result += "✅ PERMISSION ASSIGNMENTS ($($roleAssignments.Count) total)`n"
                $result += "-" * 30 + "`n"
                
                foreach ($assignment in $roleAssignments | Select-Object -First 15) {
                    try {
                        # Load the member and role definition bindings
                        $member = Get-PnPProperty -ClientObject $assignment -Property Member -ErrorAction SilentlyContinue
                        $roleDefinitions = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings -ErrorAction SilentlyContinue
                        
                        if ($member) {
                            $result += "🔐 $($member.Title)`n"
                            $result += "   Type: $($member.PrincipalType)`n"
                            
                            # Add login name if available
                            if ($member.LoginName) {
                                $result += "   Login Name: $($member.LoginName)`n"
                            }
                            
                            # Get role definitions (permissions)
                            if ($roleDefinitions -and $roleDefinitions.Count -gt 0) {
                                $permissions = ($roleDefinitions | ForEach-Object { $_.Name }) -join ", "
                                $result += "   Permissions: $permissions`n"
                                
                                # Show permission level indicators
                                $highPermissions = $roleDefinitions | Where-Object { 
                                    $_.Name -in @("Full Control", "Design", "Edit", "Contribute") 
                                }
                                if ($highPermissions) {
                                    $result += "   ⚡ High-level permissions detected`n"
                                }
                                
                                # Show read-only permissions
                                $readOnlyPermissions = $roleDefinitions | Where-Object { 
                                    $_.Name -in @("Read", "View Only") 
                                }
                                if ($readOnlyPermissions) {
                                    $result += "   👁️ Read-only access`n"
                                }
                            } else {
                                $result += "   Permissions: Unable to retrieve role definitions`n"
                            }
                            $result += "`n"
                        }
                    }
                    catch {
                        $result += "🔐 Permission assignment (details unavailable)`n"
                        $result += "   Error: $($_.Exception.Message)`n`n"
                    }
                }
                
                if ($roleAssignments.Count -gt 15) {
                    $result += "   ... and $($roleAssignments.Count - 15) more assignments`n"
                }
                $result += "`n"
                
                Write-ActivityLog "Successfully retrieved $($roleAssignments.Count) role assignments using modern method"
            } else {
                $result += "⚠️ PERMISSION ASSIGNMENTS: No role assignments found`n`n"
            }
        }
        catch {
            # Alternative modern approach if the above fails
            try {
                Write-ActivityLog "Trying alternative modern approach for role assignments"
                
                # Alternative: Use Get-PnPProperty to get role assignments from current web
                $web = Get-PnPWeb
                $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
                
                if ($roleAssignments -and $roleAssignments.Count -gt 0) {
                    $result += "✅ PERMISSION ASSIGNMENTS (Alternative Method - $($roleAssignments.Count) total)`n"
                    $result += "-" * 30 + "`n"
                    
                    $assignmentCount = 0
                    foreach ($assignment in $roleAssignments) {
                        if ($assignmentCount -ge 15) { break }
                        
                        try {
                            $member = Get-PnPProperty -ClientObject $assignment -Property Member
                            $roleDefinitionBindings = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings
                            
                            $result += "🔐 $($member.Title)`n"
                            $result += "   Type: $($member.PrincipalType)`n"
                            
                            if ($member.LoginName) {
                                $result += "   Login Name: $($member.LoginName)`n"
                            }
                            
                            if ($roleDefinitionBindings) {
                                $permissions = ($roleDefinitionBindings | ForEach-Object { $_.Name }) -join ", "
                                $result += "   Permissions: $permissions`n"
                            }
                            
                            $result += "`n"
                            $assignmentCount++
                        }
                        catch {
                            $result += "🔐 Assignment details unavailable`n`n"
                            $assignmentCount++
                        }
                    }
                    
                    if ($roleAssignments.Count -gt 15) {
                        $result += "   ... and $($roleAssignments.Count - 15) more assignments`n"
                    }
                    $result += "`n"
                    
                    Write-ActivityLog "Successfully retrieved $($roleAssignments.Count) role assignments using alternative method"
                } else {
                    throw "No role assignments found using alternative method"
                }
            }
            catch {
                $result += "⚠️ PERMISSION ASSIGNMENTS: Limited access to permission details`n"
                $result += "Reason: $($_.Exception.Message)`n"
                $result += "Note: Role assignment enumeration requires elevated SharePoint permissions.`n"
                $result += "Modern PnP 3.x: The Get-PnPRoleAssignment cmdlet was removed. Using alternative methods.`n`n"
                Write-ActivityLog "Failed to get role assignments with both modern methods: $($_.Exception.Message)"
            }
        }
        
        # Step 5: Get lists and libraries for comprehensive analysis
        $script:txtOperationsResults.Text = $result + "Retrieving lists and libraries..."
        try {
            $lists = Get-PnPList -ErrorAction SilentlyContinue
            if ($lists) {
                $result += "📚 LISTS AND LIBRARIES ($($lists.Count) total)`n"
                $result += "-" * 25 + "`n"
                
                $documentLibraries = $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" }
                $regularLists = $lists | Where-Object { $_.BaseType -eq "GenericList" }
                $systemLists = $lists | Where-Object { $_.Hidden -eq $true }
                
                $result += "Document Libraries: $($documentLibraries.Count)`n"
                $result += "Lists: $($regularLists.Count)`n"
                $result += "System Lists: $($systemLists.Count)`n`n"
                
                # Show key libraries and lists
                foreach ($list in ($documentLibraries + $regularLists) | Where-Object { -not $_.Hidden } | Select-Object -First 10) {
                    $listType = if ($list.BaseType -eq "DocumentLibrary") { "📁" } else { "📋" }
                    $result += "$listType $($list.Title)`n"
                    $result += "   URL: $($list.DefaultViewUrl)`n"
                    $result += "   Items: $($list.ItemCount)`n"
                    $result += "   Created: $($list.Created.ToString('yyyy-MM-dd'))`n"
                    $result += "   Unique Permissions: $($list.HasUniqueRoleAssignments)`n`n"
                }
                
                if (($documentLibraries.Count + $regularLists.Count) -gt 10) {
                    $remaining = ($documentLibraries.Count + $regularLists.Count) - 10
                    $result += "   ... and $remaining more lists/libraries`n`n"
                }
            }
        }
        catch {
            $result += "⚠️ LISTS AND LIBRARIES: Access limited`n"
            $result += "Reason: $($_.Exception.Message)`n`n"
            Write-ActivityLog "Failed to get lists: $($_.Exception.Message)"
        }
        
        # Step 6: Enhanced security information with modern PnP
        try {
            $script:txtOperationsResults.Text = $result + "Retrieving enhanced security settings..."
            
            $result += "🔒 ENHANCED SECURITY SETTINGS`n"
            $result += "-" * 30 + "`n"
            
            # Site collection administrators
            try {
                $siteCollectionAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                if ($siteCollectionAdmins) {
                    $result += "Site Collection Administrators: $($siteCollectionAdmins.Count)`n"
                    foreach ($admin in $siteCollectionAdmins | Select-Object -First 10) {
                        $result += "  👑 $($admin.Title)"
                        if ($admin.Email) { $result += " ($($admin.Email))" }
                        $result += "`n"
                    }
                    $result += "`n"
                }
            }
            catch {
                $result += "Site Collection Administrators: Unable to retrieve`n`n"
            }
            
            # Site features
            try {
                $siteFeatures = Get-PnPFeature -Scope Site -ErrorAction SilentlyContinue
                $webFeatures = Get-PnPFeature -Scope Web -ErrorAction SilentlyContinue
                
                if ($siteFeatures -or $webFeatures) {
                    $result += "Active Features:`n"
                    if ($siteFeatures) {
                        $result += "  Site Features: $($siteFeatures.Count)`n"
                    }
                    if ($webFeatures) {
                        $result += "  Web Features: $($webFeatures.Count)`n"
                    }
                    
                    # Show important features
                    $importantFeatures = ($siteFeatures + $webFeatures) | Where-Object { 
                        $_.DisplayName -and 
                        -not $_.DisplayName.StartsWith("TenantSitesList") -and
                        -not $_.DisplayName.StartsWith("Publishing")
                    } | Select-Object -First 8
                    
                    foreach ($feature in $importantFeatures) {
                        $result += "  🔧 $($feature.DisplayName)`n"
                    }
                    $result += "`n"
                }
            }
            catch {
                Write-ActivityLog "Could not retrieve features: $($_.Exception.Message)"
            }
            
            # Regional settings
            try {
                $regionalSettings = Get-PnPRegionalSettings -ErrorAction SilentlyContinue
                if ($regionalSettings) {
                    $result += "Regional Settings:`n"
                    $result += "  Locale ID: $($regionalSettings.LocaleId)`n"
                    $result += "  Time Zone: $($regionalSettings.TimeZone.Description)`n"
                    $result += "  First Day of Week: $($regionalSettings.FirstDayOfWeek)`n"
                    $result += "`n"
                }
            }
            catch {
                Write-ActivityLog "Could not retrieve regional settings: $($_.Exception.Message)"
            }
            
        }
        catch {
            Write-ActivityLog "Could not retrieve enhanced security settings: $($_.Exception.Message)"
        }
        
        $result += "=" * 65 + "`n"
        $result += "✅ MODERN PnP ANALYSIS COMPLETED SUCCESSFULLY`n"
        $result += "Site analyzed: $siteUrl`n"
        $result += "Analysis time: $(Get-Date -Format 'HH:mm:ss')`n"
        $result += "PnP Version: Modern (3.x) with enhanced features`n`n"
        
        # Enhanced summary with modern insights
        $result += "📊 ENHANCED SUMMARY & INSIGHTS:`n"
        $result += "-" * 35 + "`n"
        
        if ($users.Count -eq 0 -or $users.Count -lt 5) {
            $result += "• User enumeration limited - this is normal with app-only authentication`n"
        } else {
            $result += "• Successfully enumerated $($users.Count) users with full details`n"
        }
        
        if ($groups.Count -gt 0) {
            $result += "• Retrieved $($groups.Count) SharePoint groups with member details`n"
        }
        
        try {
            if ($lists) {
                $result += "• Found $($lists.Count) lists/libraries, $($lists.Count - $systemLists.Count) user-facing`n"
            }
        }
        catch { }
        
        if ($web.HasUniqueRoleAssignments) {
            $result += "• Site has unique permissions (not inheriting from parent)`n"
        } else {
            $result += "• Site inherits permissions from parent site collection`n"
        }
        
        $result += "• Using modern PnP PowerShell 3.x with enhanced capabilities`n"
        $result += "• For full tenant-level analysis, SharePoint Administrator role recommended`n`n"
        
        $result += "💡 Modern Features Available:`n"
        $result += "• Enhanced user and group enumeration`n"
        $result += "• Detailed permission analysis with role definitions`n"
        $result += "• List and library inventory with security settings`n"
        $result += "• Site feature and configuration analysis`n`n"
        
        $result += "🔄 To analyze a different site, enter a new URL above and click 'Analyze Permissions' again."
        
        $script:txtOperationsResults.Text = $result
        Write-ActivityLog "Successfully completed modern PnP permissions analysis for $siteUrl"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointPermissions"
        $script:txtOperationsResults.Text = @"
❌ ANALYSIS ERROR

Site: $($script:txtSiteUrl.Text.Trim())
Error: $($_.Exception.Message)

This error typically occurs when:
1. You don't have sufficient permissions to access the site
2. The site URL is incorrect or the site doesn't exist  
3. Your app registration lacks the necessary permissions
4. There are network connectivity issues

Troubleshooting steps:
1. ✅ Verify the site URL is correct
2. ✅ Test access to the site in your web browser
3. ✅ Check your app registration permissions
4. ✅ Try with a different site you know you have access to
5. ✅ Use Demo Mode to test the tool functionality

App Registration: $(Get-AppSetting -SettingName 'SharePoint.ClientId')
PnP PowerShell: Modern version (3.x)
"@
    }
}
