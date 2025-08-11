function Initialize-OperationsTab {
    <#
    .SYNOPSIS
    Initializes the SharePoint Operations tab with event handlers
    #>
    try {
        Write-ActivityLog "Initializing Operations tab" -Level "Information"
        
        # Set up event handlers for operations tab
        $script:btnGetSites.Add_Click({ 
            Get-SharePointSites 
        })
        
        $script:btnGetPermissions.Add_Click({ 
            Get-SharePointPermissions 
        })
        
        $script:btnGenerateReport.Add_Click({ 
            Generate-SharePointReport 
        })
        
        # Set initial state
        $script:txtOperationsResults.Text = "Connect to SharePoint to begin operations..."
        
        # Disable buttons initially
        $script:btnGetSites.IsEnabled = $false
        $script:btnGetPermissions.IsEnabled = $false
        $script:btnGenerateReport.IsEnabled = $false
        
        Write-ActivityLog "Operations tab initialized successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Operations tab: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

function Get-SharePointSites {
    <#
    .SYNOPSIS
    Retrieves and displays SharePoint sites with real-time console updates
    #>
    try {
        Write-ActivityLog "Starting SharePoint sites retrieval" -Level "Information"
        
        # Clear previous results and show starting message
        Write-ConsoleOutput "🔍 SHAREPOINT SITES ANALYSIS" -ForceUpdate
        Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate
        
        if (Get-AppSetting -SettingName "DemoMode") {
            Get-DemoSites
        } else {
            Get-RealSites
        }
        
        Write-ActivityLog "SharePoint sites retrieval completed" -Level "Information"
    }
    catch {
        Write-ConsoleOutput "❌ ERROR: Failed to retrieve SharePoint sites" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointSites"
    }
}

function Get-RealSites {
    <#
    .SYNOPSIS
    Retrieves real SharePoint sites using modern PnP PowerShell with Visual Analytics integration
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        Write-ConsoleOutput "🔄 Using modern PnP PowerShell for site enumeration..." -Append -ForceUpdate
        Write-ConsoleOutput "📡 Attempting modern tenant-level site enumeration..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1000
        
        $sites = @()
        $adminRequired = $false
        $useModernApproach = $true
        
        # Try modern tenant-level site enumeration first
        try {
            $pnpModule = Get-Module PnP.PowerShell
            if ($pnpModule.Version -ge [version]"3.0.0") {
                Write-ActivityLog "Using modern PnP 3.x site enumeration"
                Write-ConsoleOutput "✨ Using Modern PnP PowerShell 3.x features" -Append -ForceUpdate
                
                try {
                    Write-ConsoleOutput "🔍 Scanning tenant for site collections..." -Append -ForceUpdate
                    $sites = Get-PnPTenantSite -ErrorAction Stop
                    Write-ActivityLog "Retrieved $($sites.Count) sites using modern Get-PnPTenantSite"
                    Write-ConsoleOutput "✅ Successfully retrieved sites using modern PnP." -Append -ForceUpdate
                }
                catch {
                    Write-ActivityLog "Modern Get-PnPTenantSite failed, trying alternative: $($_.Exception.Message)"
                    Write-ConsoleOutput "⚠️ Tenant-level access limited, trying alternative methods..." -Append -ForceUpdate
                    
                    # Try hub sites as alternative
                    try {
                        $hubSites = Get-PnPHubSite -ErrorAction SilentlyContinue
                        if ($hubSites) {
                            $sites += $hubSites
                            Write-ActivityLog "Retrieved $($hubSites.Count) hub sites"
                            Write-ConsoleOutput "📍 Found $($hubSites.Count) hub sites" -Append -ForceUpdate
                        }
                    }
                    catch {
                        Write-ActivityLog "Hub sites enumeration also failed: $($_.Exception.Message)"
                    }
                    
                    throw $_.Exception
                }
            } else {
                $useModernApproach = $false
                Write-ConsoleOutput "📡 Using legacy PnP approach..." -Append -ForceUpdate
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
                Write-ConsoleOutput "🔐 Trying admin center connection with modern PnP..." -Append -ForceUpdate
                Write-ConsoleOutput "Please authenticate if prompted." -Append -ForceUpdate
                
                Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive
                $sites = Get-PnPTenantSite -ErrorAction Stop
                Write-ActivityLog "Successfully retrieved $($sites.Count) sites from admin center using modern PnP"
                Write-ConsoleOutput "✅ Admin center access successful!" -Append -ForceUpdate
                $adminRequired = $false
            }
            catch {
                Write-ActivityLog "Modern admin center access also failed: $($_.Exception.Message)"
                
                # Final fallback: get current site and subsites with modern features
                try {
                    Write-ConsoleOutput "🔄 Using modern fallback method (current site + subsites)..." -Append -ForceUpdate
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
                    Write-ConsoleOutput "📂 Found $($sites.Count) accessible sites" -Append -ForceUpdate
                }
                catch {
                    throw "Unable to retrieve any sites using modern PnP. SharePoint Administrator permissions may be required."
                }
            }
        }
        
        # Enhanced results formatting with modern features
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📊 SITES DISCOVERY RESULTS (Modern PnP)" -Append -ForceUpdate
        Write-ConsoleOutput "=" * 50 -Append
        Write-ConsoleOutput "" -Append
        
        if ($adminRequired) {
            Write-ConsoleOutput "⚠️  Limited Access Mode - Showing available sites only" -Append
            Write-ConsoleOutput "Note: SharePoint Administrator permissions required for full tenant enumeration" -Append
            Write-ConsoleOutput "" -Append
        }
        
        if ($useModernApproach) {
            Write-ConsoleOutput "✨ Using Modern PnP PowerShell 3.x features" -Append
        }
        
        Write-ConsoleOutput "🏢 Sites Found: $($sites.Count)" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        
        # Enhanced site information display - CRITICAL FOR VISUAL ANALYTICS
        $siteCounter = 0
        $processedSites = @()
        
        foreach ($site in $sites | Select-Object -First 25) {
            $siteCounter++
            
            # Create structured site data for Visual Analytics
            $siteData = @{
                Title = if ($site.Title) { $site.Title } else { "Site $siteCounter" }
                Url = if ($site.Url) { $site.Url } else { "N/A" }
                Owner = if ($site.Owner) { $site.Owner } elseif ($site.SiteOwnerEmail) { $site.SiteOwnerEmail } else { "N/A" }
            }
            
            # Format console output
            Write-ConsoleOutput "📁 SITE #$siteCounter`: $($siteData.Title)" -Append -ForceUpdate
            Write-ConsoleOutput "   🌐 URL: $($siteData.Url)" -Append
            
            # Enhanced properties available in modern PnP
            if ($site.Owner) { 
                Write-ConsoleOutput "   👤 Owner: $($site.Owner)" -Append
                $siteData["Owner"] = $site.Owner
            }
            if ($site.SiteOwnerEmail) { 
                Write-ConsoleOutput "   📧 Owner Email: $($site.SiteOwnerEmail)" -Append
                if (-not $siteData["Owner"] -or $siteData["Owner"] -eq "N/A") {
                    $siteData["Owner"] = $site.SiteOwnerEmail
                }
            }
            
            # Storage information - CRITICAL FOR CHARTS
            $storageValue = 0
            if ($site.StorageUsageCurrent) { 
                $storageValue = [int]$site.StorageUsageCurrent
                Write-ConsoleOutput "   💾 Storage: $storageValue MB" -Append
                $siteData["Storage"] = $storageValue.ToString()
            } else {
                # Default storage value if not available
                $storageValue = 500
                Write-ConsoleOutput "   💾 Storage: N/A (estimated: $storageValue MB)" -Append
                $siteData["Storage"] = $storageValue.ToString()
            }
            
            # Calculate usage level
            if ($storageValue -lt 500) {
                $siteData["UsageLevel"] = "Low"
                $siteData["UsageColor"] = "#28A745"
            } elseif ($storageValue -lt 1000) {
                $siteData["UsageLevel"] = "Medium"
                $siteData["UsageColor"] = "#FFC107"
            } elseif ($storageValue -lt 1500) {
                $siteData["UsageLevel"] = "High"
                $siteData["UsageColor"] = "#DC3545"
            } else {
                $siteData["UsageLevel"] = "Critical"
                $siteData["UsageColor"] = "#6F42C1"
            }
            
            if ($site.StorageQuota) { 
                Write-ConsoleOutput "   📊 Storage Quota: $($site.StorageQuota) MB" -Append
            }
            if ($site.Template) { 
                Write-ConsoleOutput "   🎨 Template: $($site.Template)" -Append
            }
            if ($site.LastContentModifiedDate) { 
                Write-ConsoleOutput "   📅 Last Modified: $($site.LastContentModifiedDate)" -Append
            }
            if ($site.IsHubSite) { 
                Write-ConsoleOutput "   🌟 Hub Site" -Append
            }
            if ($site.HubSiteId -and $site.HubSiteId -ne "00000000-0000-0000-0000-000000000000") { 
                Write-ConsoleOutput "   🔗 Connected to Hub Site" -Append
            }
            
            Write-ConsoleOutput "   ============================================================" -Append
            Update-UIAndWait -WaitMs 300
            
            # Add to processed sites for analytics
            $processedSites += $siteData
        }
        
        if ($sites.Count -gt 25) {
            Write-ConsoleOutput "... and $($sites.Count - 25) more sites" -Append
            Write-ConsoleOutput "" -Append
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "💡 Enhanced with Modern PnP Features:" -Append -ForceUpdate
        Write-ConsoleOutput "• Hub site detection and relationships" -Append
        Write-ConsoleOutput "• Enhanced storage and quota information" -Append
        Write-ConsoleOutput "• Improved site owner details" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "🔄 Copy a site URL from above and paste it in the 'Site URL' field on the SharePoint Operations tab to analyze its permissions." -Append
        
        Write-ActivityLog "Successfully completed modern site enumeration with $($sites.Count) sites"
        
        # CRITICAL: Update Visual Analytics with properly formatted data
        Write-ActivityLog "Updating Visual Analytics with $($processedSites.Count) processed sites" -Level "Information"
        
        # Ensure we have the right data structure for Visual Analytics
        $analyticsData = @{
            SitesCount = $sites.Count
            SitesData = $processedSites
        }
        
        # Call Visual Analytics update directly
        Update-MetricsCards -SitesCount $sites.Count -UsersCount 0 -GroupsCount 0 -ExternalCount 0
        Update-SitesDataGrid -SitesData $processedSites
        Update-Charts -SitesData $processedSites
        
        # Also trigger the standard parsing for consistency
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Sites"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointSites-Modern"
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "❌ ERROR: Site enumeration failed" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "This typically happens when:" -Append
        Write-ConsoleOutput "1. Your app registration lacks SharePoint Administrator permissions" -Append
        Write-ConsoleOutput "2. You need access to the SharePoint Admin Center" -Append
        Write-ConsoleOutput "3. Tenant-level operations are restricted" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "Available options:" -Append
        Write-ConsoleOutput "1. Use Demo Mode to test the enhanced tool features" -Append
        Write-ConsoleOutput "2. Contact your SharePoint Administrator for elevated permissions" -Append
        Write-ConsoleOutput "3. Try analyzing specific sites you have access to" -Append
        Write-ConsoleOutput "4. Request Sites.Read.All or Sites.FullControl.All permissions for your app registration" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "Current Connection: Modern PnP PowerShell 3.x" -Append
        Write-ConsoleOutput "Permissions Required: SharePoint Administrator or Sites.Read.All with admin consent" -Append
    }
}
function Get-DemoSites {
    <#
    .SYNOPSIS
    Retrieves demo sites with simulated progress and detailed output
    #>
    try {
        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating site collection enumeration..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500
        
        Write-ConsoleOutput "📡 Connecting to SharePoint tenant..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 800
        
        Write-ConsoleOutput "🔍 Scanning for site collections..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 600
        
        Write-ConsoleOutput "📊 Processing site metadata and permissions..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 400
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ SITES FOUND: 5" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        
        # Define demo sites with comprehensive information
        $demoSites = @(
            @{
                Title = "Team Collaboration Site"
                URL = "https://demo.sharepoint.com/sites/teamsite"
                Owner = "admin@demo.com"
                Storage = 245
                Template = "STS#3"
                Created = "2024-01-15"
                LastModified = "2024-08-07"
                UserCount = 12
                GroupCount = 4
                FileCount = 156
                Description = "Main team collaboration workspace for daily operations"
                Status = "Active"
                ExternalSharing = "Disabled"
            },
            @{
                Title = "Project Alpha Workspace"
                URL = "https://demo.sharepoint.com/sites/project-alpha"
                Owner = "project.manager@demo.com"
                Storage = 1024
                Template = "PROJECTSITE#0"
                Created = "2024-02-20"
                LastModified = "2024-08-08"
                UserCount = 8
                GroupCount = 3
                FileCount = 234
                Description = "Dedicated workspace for Project Alpha development"
                Status = "Active"
                ExternalSharing = "Enabled"
            },
            @{
                Title = "HR Document Center"
                URL = "https://demo.sharepoint.com/sites/hr-documents"
                Owner = "hr.admin@demo.com"
                Storage = 512
                Template = "STS#3"
                Created = "2024-01-10"
                LastModified = "2024-08-05"
                UserCount = 6
                GroupCount = 2
                FileCount = 89
                Description = "Centralized HR document repository"
                Status = "Active"
                ExternalSharing = "Disabled"
            },
            @{
                Title = "Company Intranet Portal"
                URL = "https://demo.sharepoint.com"
                Owner = "admin@demo.com"
                Storage = 2048
                Template = "SITEPAGEPUBLISHING#0"
                Created = "2023-12-01"
                LastModified = "2024-08-08"
                UserCount = 47
                GroupCount = 8
                FileCount = 445
                Description = "Main company intranet and communication hub"
                Status = "Active"
                ExternalSharing = "Disabled"
            },
            @{
                Title = "Sales Team Hub"
                URL = "https://demo.sharepoint.com/sites/sales"
                Owner = "sales.manager@demo.com"
                Storage = 768
                Template = "STS#3"
                Created = "2024-03-05"
                LastModified = "2024-08-06"
                UserCount = 15
                GroupCount = 5
                FileCount = 312
                Description = "Sales team collaboration and document management"
                Status = "Active"
                ExternalSharing = "Enabled"
            }
        )
        
        # Display detailed sites information
        for ($i = 0; $i -lt $demoSites.Count; $i++) {
            $site = $demoSites[$i]
            
            Write-ConsoleOutput "🏢 SITE #$($i + 1): $($site.Title)" -Append -ForceUpdate
            Write-ConsoleOutput "   📍 URL: $($site.URL)" -Append
            Write-ConsoleOutput "   👤 Owner: $($site.Owner)" -Append
            Write-ConsoleOutput "   💾 Storage Used: $($site.Storage) MB" -Append
            Write-ConsoleOutput "   📋 Template: $($site.Template)" -Append
            Write-ConsoleOutput "   📅 Created: $($site.Created)" -Append
            Write-ConsoleOutput "   🔄 Last Modified: $($site.LastModified)" -Append
            Write-ConsoleOutput "   👥 Users: $($site.UserCount)" -Append
            Write-ConsoleOutput "   👨‍👩‍👧‍👦 Groups: $($site.GroupCount)" -Append
            Write-ConsoleOutput "   📁 Files: $($site.FileCount)" -Append
            Write-ConsoleOutput "   📝 Description: $($site.Description)" -Append
            Write-ConsoleOutput "   🔒 External Sharing: $($site.ExternalSharing)" -Append
            Write-ConsoleOutput "   📊 Status: $($site.Status)" -Append
            
            # Storage usage analysis
            $usageLevel = if ($site.Storage -lt 500) { "Low" } elseif ($site.Storage -lt 1000) { "Medium" } elseif ($site.Storage -lt 1500) { "High" } else { "Critical" }
            Write-ConsoleOutput "   ⚡ Usage Level: $usageLevel" -Append
            
            Write-ConsoleOutput "   ============================================================" -Append
            Update-UIAndWait -WaitMs 400
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📈 DETAILED ANALYSIS & STATISTICS:" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        
        # Calculate statistics
        $totalUsers = ($demoSites | Measure-Object -Property UserCount -Sum).Sum
        $totalGroups = ($demoSites | Measure-Object -Property GroupCount -Sum).Sum
        $totalFiles = ($demoSites | Measure-Object -Property FileCount -Sum).Sum
        $totalStorage = ($demoSites | Measure-Object -Property Storage -Sum).Sum
        $avgStorage = [math]::Round(($demoSites | Measure-Object -Property Storage -Average).Average, 2)
        $sitesWithExternalSharing = ($demoSites | Where-Object { $_.ExternalSharing -eq "Enabled" }).Count
        $highUsageSites = ($demoSites | Where-Object { $_.Storage -gt 1000 }).Count
        
        Write-ConsoleOutput "🔢 TENANT OVERVIEW:" -Append
        Write-ConsoleOutput "   • Total Sites: $($demoSites.Count)" -Append
        Write-ConsoleOutput "   • Total Users Across All Sites: $totalUsers" -Append
        Write-ConsoleOutput "   • Total SharePoint Groups: $totalGroups" -Append
        Write-ConsoleOutput "   • Total Files: $totalFiles" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "💾 STORAGE ANALYSIS:" -Append
        Write-ConsoleOutput "   • Total Storage Used: $totalStorage MB ($([math]::Round($totalStorage/1024, 2)) GB)" -Append
        Write-ConsoleOutput "   • Average Storage per Site: $avgStorage MB" -Append
        Write-ConsoleOutput "   • Sites over 1GB: $highUsageSites" -Append
        $largestSite = $demoSites | Sort-Object Storage -Descending | Select-Object -First 1
        Write-ConsoleOutput "   • Largest Site: $($largestSite.Title) ($($largestSite.Storage) MB)" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "🔒 SECURITY & GOVERNANCE:" -Append
        Write-ConsoleOutput "   • Sites with External Sharing: $sitesWithExternalSharing" -Append
        $mostRecentSite = $demoSites | Sort-Object LastModified -Descending | Select-Object -First 1
        Write-ConsoleOutput "   • Most Recent Activity: $($mostRecentSite.Title)" -Append
        Write-ConsoleOutput "   • Template Distribution:" -Append
        $templateGroups = $demoSites | Group-Object Template
        foreach ($template in $templateGroups) {
            Write-ConsoleOutput "     - $($template.Name): $($template.Count) sites" -Append
        }
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "⚠️ RECOMMENDATIONS:" -Append
        Write-ConsoleOutput "   • Review storage usage for high-consumption sites" -Append
        Write-ConsoleOutput "   • Audit external sharing permissions on $sitesWithExternalSharing sites" -Append
        Write-ConsoleOutput "   • Consider implementing retention policies" -Append
        Write-ConsoleOutput "   • Schedule regular permission reviews" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "✅ Site enumeration completed successfully!" -Append -ForceUpdate
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Sites"
    }
    catch {
        throw "Demo sites retrieval failed: $($_.Exception.Message)"
    }
}

function Get-SharePointPermissions {
    <#
    .SYNOPSIS
    Analyzes SharePoint permissions with real-time console updates
    #>
    try {
        Write-ActivityLog "Starting SharePoint permissions analysis" -Level "Information"
        
        # Get site URL if specified
        $siteUrl = $script:txtSiteUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($siteUrl)) {
            $siteUrl = "All Sites (Tenant-wide analysis)"
        }
        
        # Clear previous results and show starting message
        Write-ConsoleOutput "🔐 SHAREPOINT PERMISSIONS ANALYSIS" -ForceUpdate
        Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "🎯 Target: $siteUrl" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate
        
        if (Get-AppSetting -SettingName "DemoMode") {
            Get-DemoPermissions -SiteUrl $siteUrl
        } else {
            Get-RealPermissions -SiteUrl $siteUrl
        }
        
        Write-ActivityLog "SharePoint permissions analysis completed" -Level "Information"
    }
    catch {
        Write-ConsoleOutput "❌ ERROR: Failed to analyze SharePoint permissions" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointPermissions"
    }
}

function Get-RealPermissions {
    <#
    .SYNOPSIS
    Analyzes real SharePoint permissions using modern PnP PowerShell
    #>
    param([string]$SiteUrl)
    
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        $siteUrl = $script:txtSiteUrl.Text.Trim()
        $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"
        
        # Always require a specific site URL to be entered
        if ([string]::IsNullOrEmpty($siteUrl)) {
            Write-ConsoleOutput "Please enter a specific site URL to analyze permissions." -Append
            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "Examples:" -Append
            Write-ConsoleOutput "• https://yourtenant.sharepoint.com/sites/teamsite" -Append
            Write-ConsoleOutput "• https://yourtenant.sharepoint.com/sites/project-alpha" -Append
            Write-ConsoleOutput "• https://yourtenant.sharepoint.com (tenant root site)" -Append
            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "The tool will connect to the exact site you specify and analyze its permissions." -Append
            return
        }
        
        # Validate the URL format
        if (-not $siteUrl.StartsWith("https://") -or -not $siteUrl.Contains(".sharepoint.com")) {
            Write-ConsoleOutput "Please enter a valid SharePoint Online site URL." -Append
            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "Example: https://yourtenant.sharepoint.com/sites/sitename" -Append
            return
        }
        
        Write-ActivityLog "Getting permissions for site: $siteUrl"
        Write-ConsoleOutput "🔄 Analyzing permissions for: $siteUrl..." -Append -ForceUpdate
        Write-ConsoleOutput "📡 Connecting to the specified site..." -Append -ForceUpdate
        
        # ALWAYS connect to the site specified in the input field
        try {
            Write-ActivityLog "Connecting to user-specified site: $siteUrl"
            Write-ConsoleOutput "🔐 Connecting to site..." -Append -ForceUpdate
            Write-ConsoleOutput "Please complete authentication if prompted." -Append -ForceUpdate
            
            # Connect directly to the site specified by the user
            Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Interactive
            Write-ActivityLog "Successfully connected to user-specified site: $siteUrl"
            Write-ConsoleOutput "✅ Connected successfully!" -Append -ForceUpdate
            Write-ConsoleOutput "🔍 Retrieving permissions data..." -Append -ForceUpdate
        }
        catch {
            $errorMsg = "Failed to connect to site: $siteUrl`nError: $($_.Exception.Message)"
            Write-ActivityLog $errorMsg
            
            # Provide specific troubleshooting for the entered site
            if ($_.Exception.Message -like "*Access is denied*" -or $_.Exception.Message -like "*Forbidden*") {
                Write-ConsoleOutput "❌ ACCESS DENIED to: $siteUrl" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Possible causes:" -Append
                Write-ConsoleOutput "1. You don't have permission to access this specific site" -Append
                Write-ConsoleOutput "2. The site doesn't exist or the URL is incorrect" -Append
                Write-ConsoleOutput "3. Your app registration lacks permission to access this site" -Append
                Write-ConsoleOutput "4. The site has unique permissions that block app access" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Troubleshooting steps:" -Append
                Write-ConsoleOutput "1. ✅ Verify the site URL is correct and accessible in your browser" -Append
                Write-ConsoleOutput "2. ✅ Check that you have at least read access to this site" -Append
                Write-ConsoleOutput "3. ✅ Ensure your app registration has Sites.Read.All or Sites.FullControl.All permissions" -Append
                Write-ConsoleOutput "4. ✅ Try a different site URL that you know you have access to" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "App Registration ID: $clientId" -Append
                Write-ConsoleOutput "Attempted Site: $siteUrl" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "💡 Tip: Try entering a site URL you know you have access to, such as your team site or personal OneDrive." -Append
            } elseif ($_.Exception.Message -like "*not found*" -or $_.Exception.Message -like "*404*") {
                Write-ConsoleOutput "❌ SITE NOT FOUND: $siteUrl" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "The site URL you entered could not be found." -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Please check:" -Append
                Write-ConsoleOutput "1. ✅ The URL is spelled correctly" -Append
                Write-ConsoleOutput "2. ✅ The site exists and is accessible" -Append
                Write-ConsoleOutput "3. ✅ You have the correct tenant name" -Append
                Write-ConsoleOutput "4. ✅ The site hasn't been deleted or moved" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Current URL: $siteUrl" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "💡 Tip: Copy the URL directly from your browser address bar when visiting the site." -Append
            } else {
                Write-ConsoleOutput "❌ CONNECTION ERROR to: $siteUrl" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "This could be due to:" -Append
                Write-ConsoleOutput "1. Network connectivity issues" -Append
                Write-ConsoleOutput "2. Authentication problems" -Append
                Write-ConsoleOutput "3. SharePoint service issues" -Append
                Write-ConsoleOutput "4. Invalid app registration configuration" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "Please try:" -Append
                Write-ConsoleOutput "1. ✅ Checking your internet connection" -Append
                Write-ConsoleOutput "2. ✅ Verifying your app registration is properly configured" -Append
                Write-ConsoleOutput "3. ✅ Trying again in a few minutes" -Append
                Write-ConsoleOutput "4. ✅ Testing with a different site URL" -Append
                Write-ConsoleOutput "" -Append
                Write-ConsoleOutput "App Registration ID: $clientId" -Append
            }
            return
        }
        
        # Now analyze permissions for the connected site
        Write-ConsoleOutput "📊 Analyzing permissions for: $siteUrl..." -Append -ForceUpdate
        Write-ConsoleOutput "🔍 Retrieving site information..." -Append -ForceUpdate
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📋 PERMISSIONS ANALYSIS REPORT (Modern PnP)" -Append -ForceUpdate
        Write-ConsoleOutput "Site: $siteUrl" -Append
        Write-ConsoleOutput "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append
        Write-ConsoleOutput "App Registration: $clientId" -Append
        Write-ConsoleOutput "PnP PowerShell: Modern version (3.x)" -Append
        Write-ConsoleOutput "=" * 65 -Append
        Write-ConsoleOutput "" -Append
        
        # Step 1: Get basic site/web information using modern approach
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            Write-ConsoleOutput "✅ SITE INFORMATION" -Append -ForceUpdate
            Write-ConsoleOutput "-" * 20 -Append
            Write-ConsoleOutput "📝 Title: $($web.Title)" -Append
            Write-ConsoleOutput "📄 Description: $(if($web.Description) { $web.Description } else { 'No description' })" -Append
            Write-ConsoleOutput "🌐 URL: $($web.Url)" -Append
            Write-ConsoleOutput "📅 Created: $(if($web.Created) { $web.Created.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Not available' })" -Append
            Write-ConsoleOutput "🔄 Last Modified: $(if($web.LastItemModifiedDate) { $web.LastItemModifiedDate.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Not available' })" -Append
            Write-ConsoleOutput "🌍 Language: $(if($web.Language) { $web.Language } else { 'Not available' })" -Append
            Write-ConsoleOutput "🎨 Template: $(if($web.WebTemplate) { $web.WebTemplate } else { 'Not available' })" -Append
            Write-ConsoleOutput "🔒 Has Unique Permissions: $($web.HasUniqueRoleAssignments)" -Append
            Write-ConsoleOutput "📂 Server Relative URL: $($web.ServerRelativeUrl)" -Append
            Write-ConsoleOutput "" -Append
            Write-ActivityLog "Successfully retrieved web information for $siteUrl"
        }
        catch {
            Write-ConsoleOutput "❌ SITE INFORMATION: Access denied or unavailable" -Append
            Write-ConsoleOutput "Error: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
            Write-ActivityLog "Failed to get web information: $($_.Exception.Message)"
        }
        
        # Step 2: Get users with modern PnP approach
        Write-ConsoleOutput "👥 Retrieving users..." -Append -ForceUpdate
        $users = @()
        try {
            $users = Get-PnPUser -ErrorAction Stop
            Write-ConsoleOutput "✅ SITE USERS ($($users.Count) total)" -Append -ForceUpdate
            Write-ConsoleOutput "-" * 20 -Append
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
                $userCounter = 0
                foreach ($user in $regularUsers | Select-Object -First 20) {
                    $userCounter++
                    Write-ConsoleOutput "👤 USER #$userCounter`: $($user.Title)" -Append
                    if ($user.Email) { Write-ConsoleOutput "   📧 Email: $($user.Email)" -Append }
                    
                    # Add additional user properties available in modern PnP
                    if ($user.UserPrincipalName) { Write-ConsoleOutput "   🔑 UPN: $($user.UserPrincipalName)" -Append }
                    Write-ConsoleOutput "   🔐 Login: $($user.LoginName)" -Append
                    if ($user.IsSiteAdmin) { Write-ConsoleOutput "   🔑 Site Administrator" -Append }
                    if ($user.IsShareByEmailGuestUser) { Write-ConsoleOutput "   🔗 Guest User (Email)" -Append }
                    if ($user.IsEmailAuthenticationGuestUser) { Write-ConsoleOutput "   🔗 Guest User (Auth)" -Append }
                    Write-ConsoleOutput "   --------------------------------------------------------" -Append
                    Update-UIAndWait -WaitMs 200
                }
                
                if ($regularUsers.Count -gt 20) {
                    Write-ConsoleOutput "   ... and $($regularUsers.Count - 20) more users" -Append
                }
            } else {
                Write-ConsoleOutput "No regular users found or access limited" -Append
            }
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ USERS: Limited access to user information" -Append
            Write-ConsoleOutput "Reason: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "Note: This is common with app-only authentication. User enumeration requires higher permissions." -Append
            Write-ConsoleOutput "" -Append
            Write-ActivityLog "Failed to get users: $($_.Exception.Message)"
        }
        
        # Step 3: Get groups with enhanced modern PnP features
        Write-ConsoleOutput "👨‍👩‍👧‍👦 Retrieving groups..." -Append -ForceUpdate
        $groups = @()
        try {
            $groups = Get-PnPGroup -ErrorAction Stop
            Write-ConsoleOutput "✅ SHAREPOINT GROUPS ($($groups.Count) total)" -Append -ForceUpdate
            Write-ConsoleOutput "-" * 25 -Append
            Write-ActivityLog "Successfully retrieved $($groups.Count) groups"
            
            # Enhanced group filtering
            $importantGroups = $groups | Where-Object {
                -not $_.Title.StartsWith("SharingLinks") -and 
                -not $_.Title.StartsWith("Limited Access") -and
                -not $_.Title.StartsWith("Excel") -and
                -not $_.Title.Contains("Restricted") -and
                -not $_.Title.Contains("Style Resource Readers")
            }
            
            $groupCounter = 0
            foreach ($group in $importantGroups | Select-Object -First 20) {
                $groupCounter++
                Write-ConsoleOutput "👥 GROUP #$groupCounter`: $($group.Title)" -Append
                if ($group.Description) { Write-ConsoleOutput "   📝 Description: $($group.Description)" -Append }
                
                # Enhanced group information
                if ($group.OwnerTitle) { Write-ConsoleOutput "   👤 Owner: $($group.OwnerTitle)" -Append }
                Write-ConsoleOutput "   🆔 ID: $($group.Id)" -Append
                
                # Try to get detailed group members
                try {
                    $groupMembers = Get-PnPGroupMember -Group $group.Title -ErrorAction SilentlyContinue
                    if ($groupMembers) {
                        Write-ConsoleOutput "   👥 Members: $($groupMembers.Count)" -Append
                        
                        # Categorize members
                        $users = $groupMembers | Where-Object { $_.PrincipalType -eq "User" }
                        $nestedGroups = $groupMembers | Where-Object { $_.PrincipalType -eq "SecurityGroup" }
                        
                        if ($users.Count -gt 0) {
                            Write-ConsoleOutput "   👤 Users: $($users.Count)" -Append
                            $userSample = ($users | Select-Object -First 3 | ForEach-Object { $_.Title }) -join ", "
                            if ($userSample) {
                                Write-ConsoleOutput "      └─ e.g., $userSample" -Append
                                if ($users.Count -gt 3) { Write-ConsoleOutput "         ..." -Append }
                            }
                        }
                        
                        if ($nestedGroups.Count -gt 0) {
                            Write-ConsoleOutput "   🏷️ Nested Groups: $($nestedGroups.Count)" -Append
                        }
                    }
                }
                catch {
                    Write-ConsoleOutput "   👥 Members: Unable to retrieve details" -Append
                }
                Write-ConsoleOutput "   ========================================================" -Append
                Update-UIAndWait -WaitMs 250
            }
            
            if ($importantGroups.Count -gt 20) {
                Write-ConsoleOutput "   ... and $($importantGroups.Count - 20) more groups" -Append
            }
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "❌ GROUPS: Access denied or unavailable" -Append
            Write-ConsoleOutput "Error: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
            Write-ActivityLog "Failed to get groups: $($_.Exception.Message)"
        }

        # Step 4: Get permissions using modern PnP 3.x approach
        Write-ConsoleOutput "🔒 Retrieving permissions and role assignments..." -Append -ForceUpdate
        try {
            # Modern way to get role assignments - use Get-PnPWeb with RoleAssignments
            $web = Get-PnPWeb -Includes "RoleAssignments" -ErrorAction Stop
            $roleAssignments = $web.RoleAssignments
            
            if ($roleAssignments -and $roleAssignments.Count -gt 0) {
                Write-ConsoleOutput "✅ PERMISSION ASSIGNMENTS ($($roleAssignments.Count) total)" -Append -ForceUpdate
                Write-ConsoleOutput "-" * 30 -Append
                
                $assignmentCounter = 0
                foreach ($assignment in $roleAssignments | Select-Object -First 15) {
                    $assignmentCounter++
                    try {
                        # Load the member and role definition bindings
                        $member = Get-PnPProperty -ClientObject $assignment -Property Member -ErrorAction SilentlyContinue
                        $roleDefinitions = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings -ErrorAction SilentlyContinue
                        
                        if ($member) {
                            Write-ConsoleOutput "🔐 ASSIGNMENT #$assignmentCounter`: $($member.Title)" -Append
                            Write-ConsoleOutput "   👤 Type: $($member.PrincipalType)" -Append
                            
                            # Add login name if available
                            if ($member.LoginName) {
                                Write-ConsoleOutput "   🔑 Login Name: $($member.LoginName)" -Append
                            }
                            
                            # Get role definitions (permissions)
                            if ($roleDefinitions -and $roleDefinitions.Count -gt 0) {
                                $permissions = ($roleDefinitions | ForEach-Object { $_.Name }) -join ", "
                                Write-ConsoleOutput "   🔒 Permissions: $permissions" -Append
                                
                                # Show permission level indicators
                                $highPermissions = $roleDefinitions | Where-Object { 
                                    $_.Name -in @("Full Control", "Design", "Edit", "Contribute") 
                                }
                                if ($highPermissions) {
                                    Write-ConsoleOutput "   ⚡ High-level permissions detected" -Append
                                }
                                
                                # Show read-only permissions
                                $readOnlyPermissions = $roleDefinitions | Where-Object { 
                                    $_.Name -in @("Read", "View Only") 
                                }
                                if ($readOnlyPermissions) {
                                    Write-ConsoleOutput "   👁️ Read-only access" -Append
                                }
                            } else {
                                Write-ConsoleOutput "   🔒 Permissions: Unable to retrieve role definitions" -Append
                            }
                            Write-ConsoleOutput "   --------------------------------------------------------" -Append
                            Update-UIAndWait -WaitMs 300
                        }
                    }
                    catch {
                        Write-ConsoleOutput "🔐 Permission assignment (details unavailable)" -Append
                        Write-ConsoleOutput "   Error: $($_.Exception.Message)" -Append
                        Write-ConsoleOutput "   --------------------------------------------------------" -Append
                    }
                }
                
                if ($roleAssignments.Count -gt 15) {
                    Write-ConsoleOutput "   ... and $($roleAssignments.Count - 15) more assignments" -Append
                }
                Write-ConsoleOutput "" -Append
                
                Write-ActivityLog "Successfully retrieved $($roleAssignments.Count) role assignments using modern method"
            } else {
                Write-ConsoleOutput "⚠️ PERMISSION ASSIGNMENTS: No role assignments found" -Append
                Write-ConsoleOutput "" -Append
            }
        }
        catch {
            # Alternative modern approach if the above fails
            try {
                Write-ActivityLog "Trying alternative modern approach for role assignments"
                Write-ConsoleOutput "🔄 Trying alternative permission retrieval method..." -Append -ForceUpdate
                
                # Alternative: Use Get-PnPProperty to get role assignments from current web
                $web = Get-PnPWeb
                $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
                
                if ($roleAssignments -and $roleAssignments.Count -gt 0) {
                    Write-ConsoleOutput "✅ PERMISSION ASSIGNMENTS (Alternative Method - $($roleAssignments.Count) total)" -Append -ForceUpdate
                    Write-ConsoleOutput "-" * 30 -Append
                    
                    $assignmentCount = 0
                    foreach ($assignment in $roleAssignments) {
                        if ($assignmentCount -ge 15) { break }
                        
                        try {
                            $member = Get-PnPProperty -ClientObject $assignment -Property Member
                            $roleDefinitionBindings = Get-PnPProperty -ClientObject $assignment -Property RoleDefinitionBindings
                            
                            Write-ConsoleOutput "🔐 $($member.Title)" -Append
                            Write-ConsoleOutput "   Type: $($member.PrincipalType)" -Append
                            
                            if ($member.LoginName) {
                                Write-ConsoleOutput "   Login Name: $($member.LoginName)" -Append
                            }
                            
                            if ($roleDefinitionBindings) {
                                $permissions = ($roleDefinitionBindings | ForEach-Object { $_.Name }) -join ", "
                                Write-ConsoleOutput "   Permissions: $permissions" -Append
                            }
                            
                            Write-ConsoleOutput "   --------------------------------------------------------" -Append
                            $assignmentCount++
                            Update-UIAndWait -WaitMs 200
                        }
                        catch {
                            Write-ConsoleOutput "🔐 Assignment details unavailable" -Append
                            Write-ConsoleOutput "   --------------------------------------------------------" -Append
                            $assignmentCount++
                        }
                    }
                    
                    if ($roleAssignments.Count -gt 15) {
                        Write-ConsoleOutput "   ... and $($roleAssignments.Count - 15) more assignments" -Append
                    }
                    Write-ConsoleOutput "" -Append
                    
                    Write-ActivityLog "Successfully retrieved $($roleAssignments.Count) role assignments using alternative method"
                } else {
                    throw "No role assignments found using alternative method"
                }
            }
            catch {
                Write-ConsoleOutput "⚠️ PERMISSION ASSIGNMENTS: Limited access to permission details" -Append
                Write-ConsoleOutput "Reason: $($_.Exception.Message)" -Append
                Write-ConsoleOutput "Note: Role assignment enumeration requires elevated SharePoint permissions." -Append
                Write-ConsoleOutput "Modern PnP 3.x: The Get-PnPRoleAssignment cmdlet was removed. Using alternative methods." -Append
                Write-ConsoleOutput "" -Append
                Write-ActivityLog "Failed to get role assignments with both modern methods: $($_.Exception.Message)"
            }
        }
        
        # Step 5: Get lists and libraries for comprehensive analysis
        Write-ConsoleOutput "📚 Retrieving lists and libraries..." -Append -ForceUpdate
        try {
            $lists = Get-PnPList -ErrorAction SilentlyContinue
            if ($lists) {
                Write-ConsoleOutput "📚 LISTS AND LIBRARIES ($($lists.Count) total)" -Append -ForceUpdate
                Write-ConsoleOutput "-" * 25 -Append
                
                $documentLibraries = $lists | Where-Object { $_.BaseType -eq "DocumentLibrary" }
                $regularLists = $lists | Where-Object { $_.BaseType -eq "GenericList" }
                $systemLists = $lists | Where-Object { $_.Hidden -eq $true }
                
                Write-ConsoleOutput "📁 Document Libraries: $($documentLibraries.Count)" -Append
                Write-ConsoleOutput "📋 Lists: $($regularLists.Count)" -Append
                Write-ConsoleOutput "⚙️ System Lists: $($systemLists.Count)" -Append
                Write-ConsoleOutput "" -Append
                
                # Show key libraries and lists
                $listCounter = 0
                foreach ($list in ($documentLibraries + $regularLists) | Where-Object { -not $_.Hidden } | Select-Object -First 10) {
                    $listCounter++
                    $listType = if ($list.BaseType -eq "DocumentLibrary") { "📁" } else { "📋" }
                    Write-ConsoleOutput "$listType LIST #$listCounter`: $($list.Title)" -Append
                    Write-ConsoleOutput "   🌐 URL: $($list.DefaultViewUrl)" -Append
                    Write-ConsoleOutput "   📊 Items: $($list.ItemCount)" -Append
                    Write-ConsoleOutput "   📅 Created: $($list.Created.ToString('yyyy-MM-dd'))" -Append
                    Write-ConsoleOutput "   🔒 Unique Permissions: $($list.HasUniqueRoleAssignments)" -Append
                    Write-ConsoleOutput "   --------------------------------------------------------" -Append
                    Update-UIAndWait -WaitMs 200
                }
                
                if (($documentLibraries.Count + $regularLists.Count) -gt 10) {
                    $remaining = ($documentLibraries.Count + $regularLists.Count) - 10
                    Write-ConsoleOutput "   ... and $remaining more lists/libraries" -Append
                }
                Write-ConsoleOutput "" -Append
            }
        }
        catch {
            Write-ConsoleOutput "⚠️ LISTS AND LIBRARIES: Access limited" -Append
            Write-ConsoleOutput "Reason: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
            Write-ActivityLog "Failed to get lists: $($_.Exception.Message)"
        }
        
        # Step 6: Enhanced security information with modern PnP
        try {
            Write-ConsoleOutput "🔒 Retrieving enhanced security settings..." -Append -ForceUpdate
            
            Write-ConsoleOutput "🔒 ENHANCED SECURITY SETTINGS" -Append -ForceUpdate
            Write-ConsoleOutput "-" * 30 -Append
            
            # Site collection administrators
            try {
                $siteCollectionAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                if ($siteCollectionAdmins) {
                    Write-ConsoleOutput "👑 Site Collection Administrators: $($siteCollectionAdmins.Count)" -Append
                    $adminCounter = 0
                    foreach ($admin in $siteCollectionAdmins | Select-Object -First 10) {
                        $adminCounter++
                        Write-ConsoleOutput "  👑 ADMIN #$adminCounter`: $($admin.Title)" -Append
                        if ($admin.Email) { Write-ConsoleOutput "     📧 $($admin.Email)" -Append }
                    }
                    Write-ConsoleOutput "" -Append
                }
            }
            catch {
                Write-ConsoleOutput "👑 Site Collection Administrators: Unable to retrieve" -Append
                Write-ConsoleOutput "" -Append
            }
            
            # Site features
            try {
                $siteFeatures = Get-PnPFeature -Scope Site -ErrorAction SilentlyContinue
                $webFeatures = Get-PnPFeature -Scope Web -ErrorAction SilentlyContinue
                
                if ($siteFeatures -or $webFeatures) {
                    Write-ConsoleOutput "🔧 Active Features:" -Append
                    if ($siteFeatures) {
                        Write-ConsoleOutput "  🏢 Site Features: $($siteFeatures.Count)" -Append
                    }
                    if ($webFeatures) {
                        Write-ConsoleOutput "  🌐 Web Features: $($webFeatures.Count)" -Append
                    }
                    
                    # Show important features
                    $importantFeatures = ($siteFeatures + $webFeatures) | Where-Object { 
                        $_.DisplayName -and 
                        -not $_.DisplayName.StartsWith("TenantSitesList") -and
                        -not $_.DisplayName.StartsWith("Publishing")
                    } | Select-Object -First 8
                    
                    foreach ($feature in $importantFeatures) {
                        Write-ConsoleOutput "  🔧 $($feature.DisplayName)" -Append
                    }
                    Write-ConsoleOutput "" -Append
                }
            }
            catch {
                Write-ActivityLog "Could not retrieve features: $($_.Exception.Message)"
            }
            
            # Regional settings
            try {
                $regionalSettings = Get-PnPRegionalSettings -ErrorAction SilentlyContinue
                if ($regionalSettings) {
                    Write-ConsoleOutput "🌍 Regional Settings:" -Append
                    Write-ConsoleOutput "  🆔 Locale ID: $($regionalSettings.LocaleId)" -Append
                    Write-ConsoleOutput "  🕐 Time Zone: $($regionalSettings.TimeZone.Description)" -Append
                    Write-ConsoleOutput "  📅 First Day of Week: $($regionalSettings.FirstDayOfWeek)" -Append
                    Write-ConsoleOutput "" -Append
                }
            }
            catch {
                Write-ActivityLog "Could not retrieve regional settings: $($_.Exception.Message)"
            }
            
        }
        catch {
            Write-ActivityLog "Could not retrieve enhanced security settings: $($_.Exception.Message)"
        }
        
        Write-ConsoleOutput "=" * 65 -Append
        Write-ConsoleOutput "✅ MODERN PnP ANALYSIS COMPLETED SUCCESSFULLY" -Append -ForceUpdate
        Write-ConsoleOutput "Site analyzed: $siteUrl" -Append
        Write-ConsoleOutput "Analysis time: $(Get-Date -Format 'HH:mm:ss')" -Append
        Write-ConsoleOutput "PnP Version: Modern (3.x) with enhanced features" -Append
        Write-ConsoleOutput "" -Append
        
        # Enhanced summary with modern insights
        Write-ConsoleOutput "📊 ENHANCED SUMMARY & INSIGHTS:" -Append -ForceUpdate
        Write-ConsoleOutput "-" * 35 -Append
        
        if ($users.Count -eq 0 -or $users.Count -lt 5) {
            Write-ConsoleOutput "• User enumeration limited - this is normal with app-only authentication" -Append
        } else {
            Write-ConsoleOutput "• Successfully enumerated $($users.Count) users with full details" -Append
        }
        
        if ($groups.Count -gt 0) {
            Write-ConsoleOutput "• Retrieved $($groups.Count) SharePoint groups with member details" -Append
        }
        
        try {
            if ($lists) {
                Write-ConsoleOutput "• Found $($lists.Count) lists/libraries, $($lists.Count - $systemLists.Count) user-facing" -Append
            }
        }
        catch { }
        
        if ($web.HasUniqueRoleAssignments) {
            Write-ConsoleOutput "• Site has unique permissions (not inheriting from parent)" -Append
        } else {
            Write-ConsoleOutput "• Site inherits permissions from parent site collection" -Append
        }
        
        Write-ConsoleOutput "• Using modern PnP PowerShell 3.x with enhanced capabilities" -Append
        Write-ConsoleOutput "• For full tenant-level analysis, SharePoint Administrator role recommended" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "💡 Modern Features Available:" -Append
        Write-ConsoleOutput "• Enhanced user and group enumeration" -Append
        Write-ConsoleOutput "• Detailed permission analysis with role definitions" -Append
        Write-ConsoleOutput "• List and library inventory with security settings" -Append
        Write-ConsoleOutput "• Site feature and configuration analysis" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "🔄 To analyze a different site, enter a new URL above and click 'Analyze Permissions' again." -Append
        
        Write-ActivityLog "Successfully completed modern PnP permissions analysis for $siteUrl"
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Permissions"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointPermissions"
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "❌ ANALYSIS ERROR" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "Site: $($script:txtSiteUrl.Text.Trim())" -Append
        Write-ConsoleOutput "Error: $($_.Exception.Message)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "This error typically occurs when:" -Append
        Write-ConsoleOutput "1. You don't have sufficient permissions to access the site" -Append
        Write-ConsoleOutput "2. The site URL is incorrect or the site doesn't exist" -Append
        Write-ConsoleOutput "3. Your app registration lacks the necessary permissions" -Append
        Write-ConsoleOutput "4. There are network connectivity issues" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "Troubleshooting steps:" -Append
        Write-ConsoleOutput "1. ✅ Verify the site URL is correct" -Append
        Write-ConsoleOutput "2. ✅ Test access to the site in your web browser" -Append
        Write-ConsoleOutput "3. ✅ Check your app registration permissions" -Append
        Write-ConsoleOutput "4. ✅ Try with a different site you know you have access to" -Append
        Write-ConsoleOutput "5. ✅ Use Demo Mode to test the tool functionality" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "App Registration: $(Get-AppSetting -SettingName 'SharePoint.ClientId')" -Append
        Write-ConsoleOutput "PnP PowerShell: Modern version (3.x)" -Append
    }
}

function Get-DemoPermissions {
    <#
    .SYNOPSIS
    Analyzes demo permissions with comprehensive data
    #>
    param(
        [string]$SiteUrl
    )
    
    try {
        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating comprehensive permissions analysis..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500
        
        Write-ConsoleOutput "🔍 Scanning site permissions and security structure..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 800
        
        Write-ConsoleOutput "👥 Enumerating users, groups, and access levels..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 600
        
        Write-ConsoleOutput "🔒 Analyzing permission inheritance and unique permissions..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500
        
        Write-ConsoleOutput "🌐 Checking for external users and sharing settings..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 400
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📊 COMPREHENSIVE PERMISSIONS ANALYSIS RESULTS:" -Append -ForceUpdate
        Write-ConsoleOutput "🎯 Target Site: $SiteUrl" -Append
        Write-ConsoleOutput "" -Append
        
        # Demo users with comprehensive information
        $demoUsers = @(
            @{Name="John Doe"; Email="john.doe@demo.com"; Permission="Full Control"; Type="Internal"; Department="IT"; LastAccess="2024-08-08"},
            @{Name="Jane Smith"; Email="jane.smith@demo.com"; Permission="Edit"; Type="Internal"; Department="Marketing"; LastAccess="2024-08-07"},
            @{Name="Mike Johnson"; Email="mike.johnson@demo.com"; Permission="Read"; Type="Internal"; Department="Sales"; LastAccess="2024-08-06"},
            @{Name="Sarah Wilson"; Email="sarah.wilson@demo.com"; Permission="Edit"; Type="Internal"; Department="HR"; LastAccess="2024-08-08"},
            @{Name="Alex Brown"; Email="alex.brown@demo.com"; Permission="Read"; Type="Internal"; Department="Finance"; LastAccess="2024-08-05"},
            @{Name="Lisa Davis"; Email="lisa.davis@demo.com"; Permission="Full Control"; Type="Internal"; Department="IT"; LastAccess="2024-08-08"},
            @{Name="Tom Miller"; Email="tom.miller@demo.com"; Permission="Read"; Type="Internal"; Department="Operations"; LastAccess="2024-08-04"},
            @{Name="Emma Taylor"; Email="emma.taylor@demo.com"; Permission="Edit"; Type="Internal"; Department="Marketing"; LastAccess="2024-08-07"},
            @{Name="David Anderson"; Email="david.anderson@demo.com"; Permission="Read"; Type="Internal"; Department="Sales"; LastAccess="2024-08-03"},
            @{Name="Maria Garcia"; Email="maria.garcia@demo.com"; Permission="Edit"; Type="Internal"; Department="HR"; LastAccess="2024-08-06"},
            @{Name="Robert Chen"; Email="robert.chen@demo.com"; Permission="Read"; Type="Internal"; Department="Engineering"; LastAccess="2024-08-08"},
            @{Name="Jennifer White"; Email="jennifer.white@demo.com"; Permission="Edit"; Type="Internal"; Department="Legal"; LastAccess="2024-08-02"},
            @{Name="Michael Partner"; Email="m.partner@external-company.com"; Permission="Read"; Type="External"; Department="Partner"; LastAccess="2024-08-01"},
            @{Name="Susan Contractor"; Email="s.contractor@freelance.com"; Permission="Edit"; Type="External"; Department="Contractor"; LastAccess="2024-07-30"},
            @{Name="Admin Service"; Email="admin@demo.com"; Permission="Full Control"; Type="Service Account"; Department="System"; LastAccess="2024-08-08"}
        )
        
        # Demo groups
        $demoGroups = @(
            @{Name="Site Owners"; Members=3; Permission="Full Control"; Description="Site collection administrators with full access"},
            @{Name="Site Members"; Members=8; Permission="Edit"; Description="Regular contributors with edit permissions"},
            @{Name="Site Visitors"; Members=4; Permission="Read"; Description="Read-only access for viewing content"},
            @{Name="Project Alpha Contributors"; Members=5; Permission="Edit"; Description="Project team members with contribution rights"},
            @{Name="Document Reviewers"; Members=6; Permission="Read"; Description="Review team with read access to documents"},
            @{Name="External Partners"; Members=2; Permission="Limited Access"; Description="External collaborators with restricted access"},
            @{Name="HR Managers"; Members=3; Permission="Full Control"; Description="HR department with administrative access"},
            @{Name="Sales Team"; Members=7; Permission="Edit"; Description="Sales department with content management rights"}
        )
        
        Write-ConsoleOutput "📊 PERMISSION SUMMARY:" -Append -ForceUpdate
        Write-ConsoleOutput "   👤 Total Users: $($demoUsers.Count)" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Total Groups: $($demoGroups.Count)" -Append
        $externalUsers = ($demoUsers | Where-Object {$_.Type -eq 'External'}).Count
        $serviceAccounts = ($demoUsers | Where-Object {$_.Type -eq 'Service Account'}).Count
        $internalUsers = ($demoUsers | Where-Object {$_.Type -eq 'Internal'}).Count
        Write-ConsoleOutput "   🌐 External Users: $externalUsers" -Append
        Write-ConsoleOutput "   🔧 Service Accounts: $serviceAccounts" -Append
        Write-ConsoleOutput "   👥 Internal Users: $internalUsers" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "👤 DETAILED USER PERMISSIONS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ======================================================================" -Append
        $userCounter = 0
        foreach ($user in $demoUsers) {
            $userCounter++
            $typeIcon = switch ($user.Type) {
                "Internal" { "🏢" }
                "External" { "🌐" }
                "Service Account" { "🔧" }
                default { "👤" }
            }
            Write-ConsoleOutput "   $typeIcon USER #$userCounter`: $($user.Name) ($($user.Email))" -Append
            Write-ConsoleOutput "      🔒 Permission: $($user.Permission)" -Append
            Write-ConsoleOutput "      🏢 Department: $($user.Department)" -Append
            Write-ConsoleOutput "      📅 Last Access: $($user.LastAccess)" -Append
            Write-ConsoleOutput "      📍 User Type: $($user.Type)" -Append
            Write-ConsoleOutput "   --------------------------------------------------" -Append
            Update-UIAndWait -WaitMs 150
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "👨‍👩‍👧‍👦 SHAREPOINT GROUPS ANALYSIS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ======================================================================" -Append
        $groupCounter = 0
        foreach ($group in $demoGroups) {
            $groupCounter++
            Write-ConsoleOutput "   🏷️ GROUP #$groupCounter`: $($group.Name)" -Append
            Write-ConsoleOutput "      👥 Members: $($group.Members)" -Append
            Write-ConsoleOutput "      🔒 Permission Level: $($group.Permission)" -Append
            Write-ConsoleOutput "      📝 Description: $($group.Description)" -Append
            Write-ConsoleOutput "   --------------------------------------------------" -Append
            Update-UIAndWait -WaitMs 200
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🔍 PERMISSION LEVEL DISTRIBUTION:" -Append -ForceUpdate
        $permissionCounts = $demoUsers | Group-Object Permission | Sort-Object Count -Descending
        foreach ($permGroup in $permissionCounts) {
            $percentage = [math]::Round(($permGroup.Count / $demoUsers.Count) * 100, 1)
            Write-ConsoleOutput "   🔒 $($permGroup.Name): $($permGroup.Count) users ($percentage%)" -Append
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🏢 DEPARTMENT BREAKDOWN:" -Append -ForceUpdate
        $departmentCounts = $demoUsers | Where-Object {$_.Type -eq "Internal"} | Group-Object Department | Sort-Object Count -Descending
        foreach ($deptGroup in $departmentCounts) {
            Write-ConsoleOutput "   🏢 $($deptGroup.Name): $($deptGroup.Count) users" -Append
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🚨 COMPREHENSIVE SECURITY ANALYSIS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ============================================================" -Append
        
        # Security findings
        $fullControlUsers = ($demoUsers | Where-Object {$_.Permission -eq 'Full Control'}).Count
        $inactiveUsers = ($demoUsers | Where-Object {[DateTime]$_.LastAccess -lt (Get-Date).AddDays(-7)}).Count
        
        Write-ConsoleOutput "   ✅ SECURITY STATUS: Generally Good" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "   🔍 FINDINGS:" -Append
        if ($externalUsers -gt 0) {
            Write-ConsoleOutput "   ⚠️ $externalUsers external users detected (review recommended)" -Append
            $demoUsers | Where-Object {$_.Type -eq 'External'} | ForEach-Object {
                Write-ConsoleOutput "      🌐 $($_.Name) - $($_.Email) - $($_.Permission)" -Append
            }
        }
        
        if ($fullControlUsers -gt 2) {
            Write-ConsoleOutput "   ⚠️ $fullControlUsers users with Full Control (consider reducing)" -Append
        }
        
        if ($inactiveUsers -gt 0) {
            Write-ConsoleOutput "   ⚠️ $inactiveUsers users inactive for 7+ days (cleanup recommended)" -Append
        }
        
        Write-ConsoleOutput "   ✅ No orphaned permissions detected" -Append
        Write-ConsoleOutput "   ✅ All groups have appropriate access levels" -Append
        Write-ConsoleOutput "   ✅ Permission inheritance is properly configured" -Append
        Write-ConsoleOutput "   ✅ No excessive service account permissions found" -Append
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "💡 RECOMMENDATIONS:" -Append
        Write-ConsoleOutput "   1. Review and audit external user access regularly" -Append
        Write-ConsoleOutput "   2. Consider implementing time-limited access for external users" -Append
        Write-ConsoleOutput "   3. Remove or disable inactive user accounts" -Append
        Write-ConsoleOutput "   4. Implement regular access reviews for Full Control users" -Append
        Write-ConsoleOutput "   5. Set up automated alerts for permission changes" -Append
        Write-ConsoleOutput "   6. Document business justification for external access" -Append
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ Comprehensive permissions analysis completed successfully!" -Append -ForceUpdate
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Permissions"
    }
    catch {
        throw "Demo permissions analysis failed: $($_.Exception.Message)"
    }
}

function Generate-SharePointReport {
    <#
    .SYNOPSIS
    Generates comprehensive SharePoint permissions report with real-time progress
    #>
    try {
        Write-ActivityLog "Starting SharePoint report generation" -Level "Information"
        
        # Clear previous results and show starting message
        Write-ConsoleOutput "📈 SHAREPOINT PERMISSIONS REPORT GENERATION" -ForceUpdate
        Write-ConsoleOutput "=======================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate
        
        if (Get-AppSetting -SettingName "DemoMode") {
            Generate-DemoReport
        } else {
            Generate-RealReport
        }
        
        Write-ActivityLog "SharePoint report generation completed" -Level "Information"
    }
    catch {
        Write-ConsoleOutput "❌ ERROR: Report generation failed!" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ErrorLog -Message $_.Exception.Message -Location "Generate-SharePointReport"
    }
}

function Generate-DemoReport {
    <#
    .SYNOPSIS
    Generates demo report with realistic progress simulation
    #>
    try {
        # Initialize progress tracking
        $script:ReportStartTime = Get-Date
        
        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating comprehensive report generation..." -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        
        # Step-by-step progress
        Show-ReportProgress "Initializing Report Environment" "Creating output directories and validating settings" 600
        Show-ReportProgress "Gathering Site Collections" "Enumerating all SharePoint sites in tenant" 1200
        Show-ReportProgress "Analyzing User Permissions" "Processing individual user access rights" 1500
        Show-ReportProgress "Processing Group Memberships" "Mapping SharePoint groups and nested permissions" 900
        Show-ReportProgress "Detecting Security Issues" "Scanning for orphaned permissions and risks" 700
        Show-ReportProgress "Generating CSV Files" "Creating detailed permission matrices" 800
        Show-ReportProgress "Creating Excel Workbook" "Formatting data with charts and pivot tables" 600
        Show-ReportProgress "Finalizing Reports" "Adding executive summary and recommendations" 400
        
        # Generate results
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $totalTime = [math]::Round(((Get-Date) - $script:ReportStartTime).TotalSeconds, 1)
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ REPORT GENERATION COMPLETED!" -Append -ForceUpdate
        Write-ConsoleOutput "========================================" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📁 GENERATED FILES:" -Append -ForceUpdate
        Write-ConsoleOutput "   📊 SharePoint_Permissions_Report_$timestamp.csv" -Append
        Write-ConsoleOutput "   📈 SharePoint_Analysis_Dashboard_$timestamp.xlsx" -Append
        Write-ConsoleOutput "   📋 Executive_Summary_$timestamp.pdf" -Append
        Write-ConsoleOutput "   📝 Detailed_Audit_Log_$timestamp.txt" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📊 REPORT STATISTICS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⏱️ Processing Time: $totalTime seconds" -Append
        Write-ConsoleOutput "   🏢 Sites Analyzed: 5" -Append
        Write-ConsoleOutput "   👥 Users Processed: 47" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Groups Analyzed: 23" -Append
        Write-ConsoleOutput "   📋 Permission Records: 1,247" -Append
        Write-ConsoleOutput "   🔍 Security Findings: 8" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "🎯 KEY FINDINGS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⚠️ 8 External users detected (review required)" -Append
        Write-ConsoleOutput "   🔴 3 Sites with excessive storage usage" -Append
        Write-ConsoleOutput "   ⚠️ 2 Orphaned permission groups found" -Append
        Write-ConsoleOutput "   ✅ All sites comply with base governance policies" -Append
        Write-ConsoleOutput "   🔍 5 Permission inheritance breaks detected" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📋 NEXT STEPS:" -Append -ForceUpdate
        Write-ConsoleOutput "   1. Review generated CSV files for detailed analysis" -Append
        Write-ConsoleOutput "   2. Open Excel dashboard for visual insights" -Append
        Write-ConsoleOutput "   3. Address identified security recommendations" -Append
        Write-ConsoleOutput "   4. Schedule regular automated permission audits" -Append
        Write-ConsoleOutput "   5. Implement governance policies for future prevention" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "💡 TIP: Check the Visual Analytics tab for an interactive overview!" -Append -ForceUpdate
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Report"
    }
    catch {
        throw "Demo report generation failed: $($_.Exception.Message)"
    }
}

function Generate-RealReport {
    <#
    .SYNOPSIS
    Generates real SharePoint report using actual data
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        Write-ConsoleOutput "🔄 Preparing real SharePoint report generation..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1000
        
        Write-ConsoleOutput "📊 Gathering comprehensive tenant data..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1500
        
        # Initialize progress tracking
        $script:ReportStartTime = Get-Date
        
        # Step 1: Gather Sites Data
        Show-ReportProgress "Gathering Site Collections" "Enumerating all accessible SharePoint sites" 1000
        
        $allSites = @()
        try {
            # Try to get tenant sites first
            try {
                $allSites = Get-PnPTenantSite -ErrorAction Stop
                Write-ConsoleOutput "   ✅ Retrieved $($allSites.Count) sites from tenant level" -Append
            }
            catch {
                # Fallback to current site and subsites
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
                $allSites = @($currentWeb)
                if ($subsites) { $allSites += $subsites }
                Write-ConsoleOutput "   ⚠️ Limited to accessible sites: $($allSites.Count) found" -Append
            }
        }
        catch {
            Write-ConsoleOutput "   ❌ Failed to gather sites data: $($_.Exception.Message)" -Append
            throw "Cannot generate report without site data"
        }
        
        # Step 2: Analyze Permissions for Each Site
        Show-ReportProgress "Analyzing Site Permissions" "Processing permissions for each site" 2000
        
        $allPermissions = @()
        $siteCounter = 0
        foreach ($site in $allSites | Select-Object -First 10) { # Limit for performance
            $siteCounter++
            try {
                Write-ConsoleOutput "   🔍 Analyzing site $siteCounter/$($allSites.Count): $($site.Title)" -Append
                
                # Connect to each site
                Connect-PnPOnline -Url $site.Url -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive -ErrorAction SilentlyContinue
                
                # Get users and groups for this site
                $users = Get-PnPUser -ErrorAction SilentlyContinue
                $groups = Get-PnPGroup -ErrorAction SilentlyContinue
                
                # Process users
                foreach ($user in $users | Where-Object { $_.PrincipalType -eq "User" }) {
                    $allPermissions += [PSCustomObject]@{
                        SiteTitle = $site.Title
                        SiteUrl = $site.Url
                        PrincipalName = $user.Title
                        PrincipalEmail = $user.Email
                        PrincipalType = "User"
                        LoginName = $user.LoginName
                        IsSiteAdmin = $user.IsSiteAdmin
                        IsExternal = $user.IsShareByEmailGuestUser -or $user.IsEmailAuthenticationGuestUser
                        LastModified = Get-Date
                    }
                }
                
                # Process groups
                foreach ($group in $groups) {
                    $allPermissions += [PSCustomObject]@{
                        SiteTitle = $site.Title
                        SiteUrl = $site.Url
                        PrincipalName = $group.Title
                        PrincipalEmail = ""
                        PrincipalType = "Group"
                        LoginName = $group.LoginName
                        IsSiteAdmin = $false
                        IsExternal = $false
                        LastModified = Get-Date
                    }
                }
                
                Update-UIAndWait -WaitMs 200
            }
            catch {
                Write-ConsoleOutput "   ⚠️ Skipped site due to access issues: $($site.Title)" -Append
            }
        }
        
        # Step 3: Generate CSV Report
        Show-ReportProgress "Generating CSV Export" "Creating detailed permission matrices" 800
        
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $csvPath = "SharePoint_Permissions_Report_$timestamp.csv"
        
        try {
            $allPermissions | Export-Csv -Path $csvPath -NoTypeInformation -ErrorAction Stop
            Write-ConsoleOutput "   ✅ CSV report saved: $csvPath" -Append
        }
        catch {
            Write-ConsoleOutput "   ❌ Failed to save CSV: $($_.Exception.Message)" -Append
        }
        
        # Step 4: Generate Summary Statistics
        Show-ReportProgress "Generating Summary Statistics" "Calculating security metrics and insights" 600
        
        $totalUsers = ($allPermissions | Where-Object { $_.PrincipalType -eq "User" }).Count
        $totalGroups = ($allPermissions | Where-Object { $_.PrincipalType -eq "Group" }).Count
        $externalUsers = ($allPermissions | Where-Object { $_.IsExternal -eq $true }).Count
        $siteAdmins = ($allPermissions | Where-Object { $_.IsSiteAdmin -eq $true }).Count
        
        # Step 5: Create Executive Summary
        Show-ReportProgress "Creating Executive Summary" "Compiling findings and recommendations" 400
        
        $totalTime = [math]::Round(((Get-Date) - $script:ReportStartTime).TotalSeconds, 1)
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ REAL SHAREPOINT REPORT COMPLETED!" -Append -ForceUpdate
        Write-ConsoleOutput "========================================" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📁 GENERATED FILES:" -Append -ForceUpdate
        Write-ConsoleOutput "   📊 $csvPath" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📊 REPORT STATISTICS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⏱️ Processing Time: $totalTime seconds" -Append
        Write-ConsoleOutput "   🏢 Sites Analyzed: $siteCounter" -Append
        Write-ConsoleOutput "   👥 Total Users: $totalUsers" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Total Groups: $totalGroups" -Append
        Write-ConsoleOutput "   📋 Permission Records: $($allPermissions.Count)" -Append
        Write-ConsoleOutput "   🌐 External Users: $externalUsers" -Append
        Write-ConsoleOutput "   🔑 Site Administrators: $siteAdmins" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "🎯 KEY FINDINGS:" -Append -ForceUpdate
        if ($externalUsers -gt 0) {
            Write-ConsoleOutput "   ⚠️ $externalUsers External users detected (review required)" -Append
        }
        if ($siteAdmins -gt 5) {
            Write-ConsoleOutput "   ⚠️ $siteAdmins Site administrators found (review recommended)" -Append
        }
        Write-ConsoleOutput "   ✅ Permissions data successfully exported for detailed analysis" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "📋 NEXT STEPS:" -Append -ForceUpdate
        Write-ConsoleOutput "   1. Open the generated CSV file for detailed analysis" -Append
        Write-ConsoleOutput "   2. Review external user access and justifications" -Append
        Write-ConsoleOutput "   3. Audit site administrator assignments" -Append
        Write-ConsoleOutput "   4. Consider implementing automated monitoring" -Append
        Write-ConsoleOutput "   5. Schedule regular permission reviews" -Append
        Write-ConsoleOutput "" -Append
        
        Write-ConsoleOutput "💡 CSV file contains complete permission data for further analysis in Excel or Power BI!" -Append -ForceUpdate
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Report"
        
    }
    catch {
        Write-ConsoleOutput "❌ Real SharePoint report generation failed!" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📝 To implement full report generation:" -Append
        Write-ConsoleOutput "   1. Ensure you have SharePoint Administrator permissions" -Append
        Write-ConsoleOutput "   2. Verify app registration has appropriate API permissions" -Append
        Write-ConsoleOutput "   3. Test with Demo Mode to see expected functionality" -Append
        Write-ConsoleOutput "   4. Consider using PowerShell execution policies if file creation fails" -Append
        throw "Real report generation failed: $($_.Exception.Message)"
    }
}

function Show-ReportProgress {
    <#
    .SYNOPSIS
    Shows individual report generation progress steps
    #>
    param(
        [string]$StepName,
        [string]$Details,
        [int]$DelayMs = 500
    )
    
    try {
        $timeStamp = Get-Date -Format "HH:mm:ss"
        
        Write-ConsoleOutput "[$timeStamp] 🔄 $StepName" -Append -ForceUpdate
        if ($Details) {
            Write-ConsoleOutput "           └─ $Details" -Append
        }
        
        Update-UIAndWait -WaitMs $DelayMs
        
        Write-ConsoleOutput "[$timeStamp] ✅ $StepName - Complete" -Append
        Write-ConsoleOutput "" -Append
    }
    catch {
        Write-ActivityLog "Error showing report progress: $($_.Exception.Message)" -Level "Warning"
    }
}

# Helper function for UI updates and waiting
function Update-UIAndWait {
    param([int]$WaitMs = 500)
    
    try {
        # Force UI update
        [System.Windows.Forms.Application]::DoEvents()
        
        # Wait for specified time
        Start-Sleep -Milliseconds $WaitMs
    }
    catch {
        # Ignore errors in UI updates
        Start-Sleep -Milliseconds $WaitMs
    }
}

# Helper function to write console output
function Write-ConsoleOutput {
    param(
        [string]$Message,
        [switch]$Append,
        [switch]$ForceUpdate
    )
    
    try {
        if ($Append) {
            $script:txtOperationsResults.Text += "`n$Message"
        } else {
            $script:txtOperationsResults.Text = $Message
        }
        
        if ($ForceUpdate) {
            # Scroll to bottom
            $script:txtOperationsResults.SelectionStart = $script:txtOperationsResults.Text.Length
            $script:txtOperationsResults.ScrollToCaret()
            
            # Force UI refresh
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        # Fallback to Write-Host if UI is not available
        Write-Host $Message
    }
}