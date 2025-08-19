# ============================================
# SitesDeepDive.ps1 - Sites Deep Dive Window
# ============================================
# Location: Functions/UI/DeepDive/SitesDeepDive.ps1

function Show-SitesDeepDive {
    <#
    .SYNOPSIS
    Shows the Sites Deep Dive window with detailed analysis
    #>
    try {
        Write-ActivityLog "Opening Sites Deep Dive window" -Level "Information"
        
        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\SitesDeepDive.xaml"
        if (-not (Test-Path $xamlPath)) {
            throw "Deep Dive XAML file not found at: $xamlPath"
        }
        
        $xamlContent = Get-Content $xamlPath -Raw
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
        $deepDiveWindow = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls
        $controls = @{
            Window = $deepDiveWindow
            # Header
            txtSiteCount = $deepDiveWindow.FindName("txtSiteCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalStorage = $deepDiveWindow.FindName("txtTotalStorage")
            txtAvgStorage = $deepDiveWindow.FindName("txtAvgStorage")
            txtHubSites = $deepDiveWindow.FindName("txtHubSites")
            txtUniquePermissions = $deepDiveWindow.FindName("txtUniquePermissions")
            txtRecentlyModified = $deepDiveWindow.FindName("txtRecentlyModified")
            # Detailed Sites Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboTemplateFilter = $deepDiveWindow.FindName("cboTemplateFilter")
            cboStorageFilter = $deepDiveWindow.FindName("cboStorageFilter")
            dgDetailedSites = $deepDiveWindow.FindName("dgDetailedSites")
            # Storage Analysis Tab
            canvasStorageDistribution = $deepDiveWindow.FindName("canvasStorageDistribution")
            dgTopStorage = $deepDiveWindow.FindName("dgTopStorage")
            # Site Health Tab
            txtHealthySites = $deepDiveWindow.FindName("txtHealthySites")
            txtWarningSites = $deepDiveWindow.FindName("txtWarningSites")
            txtCriticalSites = $deepDiveWindow.FindName("txtCriticalSites")
            lstHealthIssues = $deepDiveWindow.FindName("lstHealthIssues")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }
        
        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-SitesDeepDiveData -Controls $controls
        })
        
        $controls.btnExport.Add_Click({
            Export-SitesDeepDiveData -Controls $controls
        })
        
        $controls.txtSearch.Add_TextChanged({
            Apply-SitesFilter -Controls $controls
        })
        
        $controls.cboTemplateFilter.Add_SelectionChanged({
            Apply-SitesFilter -Controls $controls
        })
        
        $controls.cboStorageFilter.Add_SelectionChanged({
            Apply-SitesFilter -Controls $controls
        })
        
        # Load initial data
        Load-SitesDeepDiveData -Controls $controls
        
        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-SitesDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Sites Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-SitesDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Sites Deep Dive window
    #>
    param($Controls)
    
    try {
        $Controls.txtStatus.Text = "Loading sites data..."
        
        # Get sites data from the data manager
        $sites = Get-SharePointData -DataType "Sites"
        $metrics = Get-SharePointData -DataType "Metrics"
        
        if ($sites.Count -eq 0) {
            # Try to get additional site details if connected
            if ($script:SPOConnected) {
                $Controls.txtStatus.Text = "Fetching additional site details..."
                $sites = Get-DetailedSiteInformation
            }
        }
        
        # Update header
        $Controls.txtSiteCount.Text = "Analyzing $($sites.Count) sites"
        
        # Calculate summary statistics
        $totalStorage = 0
        $hubSitesCount = 0
        $uniquePermissionsCount = 0
        $recentlyModifiedCount = 0
        $cutoffDate = (Get-Date).AddDays(-7)
        
        foreach ($site in $sites) {
            $storage = [int]($site["Storage"] -replace "[^\d]", "")
            $totalStorage += $storage
            
            if ($site["IsHubSite"] -eq $true) { $hubSitesCount++ }
            if ($site["HasUniquePermissions"] -eq $true) { $uniquePermissionsCount++ }
            
            if ($site["LastModified"]) {
                try {
                    $lastMod = [DateTime]::Parse($site["LastModified"])
                    if ($lastMod -gt $cutoffDate) { $recentlyModifiedCount++ }
                }
                catch { }
            }
        }
        
        # Update summary stats
        $totalStorageGB = [math]::Round($totalStorage / 1024, 2)
        $avgStorage = if ($sites.Count -gt 0) { [math]::Round($totalStorage / $sites.Count, 0) } else { 0 }
        
        $Controls.txtTotalStorage.Text = "$totalStorageGB GB"
        $Controls.txtAvgStorage.Text = "$avgStorage MB"
        $Controls.txtHubSites.Text = $hubSitesCount.ToString()
        $Controls.txtUniquePermissions.Text = $uniquePermissionsCount.ToString()
        $Controls.txtRecentlyModified.Text = $recentlyModifiedCount.ToString()
        
        # Load detailed sites grid
        Load-DetailedSitesGrid -Controls $Controls -Sites $sites
        
        # Load storage analysis
        Load-StorageAnalysis -Controls $Controls -Sites $sites -TotalStorage $totalStorage
        
        # Load site health analysis
        Load-SiteHealthAnalysis -Controls $Controls -Sites $sites
        
        # Update status
        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SitesDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-DetailedSitesGrid {
    <#
    .SYNOPSIS
    Loads the detailed sites data grid
    #>
    param($Controls, $Sites)
    
    try {
        $siteObjects = @()
        
        foreach ($site in $Sites) {
            $siteObjects += [PSCustomObject]@{
                Title = if ($site["Title"]) { $site["Title"] } else { "Unknown" }
                Url = if ($site["Url"]) { $site["Url"] } else { "N/A" }
                Owner = if ($site["Owner"]) { $site["Owner"] } else { "N/A" }
                Storage = if ($site["Storage"]) { $site["Storage"] } else { "0" }
                Template = if ($site["Template"]) { $site["Template"] } else { "N/A" }
                LastModified = if ($site["LastModified"]) { $site["LastModified"] } else { "N/A" }
                IsHubSite = if ($site["IsHubSite"]) { $site["IsHubSite"] } else { $false }
            }
        }
        
        $Controls.dgDetailedSites.ItemsSource = $siteObjects
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-DetailedSitesGrid"
    }
}

function Load-StorageAnalysis {
    <#
    .SYNOPSIS
    Loads storage analysis data and charts
    #>
    param($Controls, $Sites, $TotalStorage)
    
    try {
        # Sort sites by storage (descending)
        $sortedSites = $Sites | Sort-Object { [int]($_["Storage"] -replace "[^\d]", "") } -Descending
        
        # Create top storage consumers list
        $topStorageData = @()
        $rank = 1
        
        foreach ($site in ($sortedSites | Select-Object -First 10)) {
            $storage = [int]($site["Storage"] -replace "[^\d]", "")
            $percentage = if ($TotalStorage -gt 0) { [math]::Round(($storage / $TotalStorage) * 100, 1) } else { 0 }
            
            $barColor = switch ($true) {
                ($storage -gt 2000) { "#DC3545" }
                ($storage -gt 1000) { "#FFC107" }
                ($storage -gt 500) { "#17A2B8" }
                default { "#28A745" }
            }
            
            $topStorageData += [PSCustomObject]@{
                Rank = $rank
                Title = $site["Title"]
                Storage = $storage
                Percentage = "$percentage%"
                PercentageValue = $percentage
                BarColor = $barColor
            }
            $rank++
        }
        
        $Controls.dgTopStorage.ItemsSource = $topStorageData
        
        # Draw storage distribution chart
        Draw-StorageDistributionChart -Canvas $Controls.canvasStorageDistribution -Sites $sortedSites
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-StorageAnalysis"
    }
}

function Draw-StorageDistributionChart {
    <#
    .SYNOPSIS
    Draws the storage distribution chart
    #>
    param($Canvas, $Sites)
    
    try {
        $Canvas.Children.Clear()
        
        if ($Sites.Count -eq 0) { return }
        
        # Take top 5 sites for visualization
        $topSites = $Sites | Select-Object -First 5
        
        $chartWidth = 400
        $chartHeight = 180
        $barWidth = 60
        $spacing = 10
        $startX = 20
        
        # Find max storage for scaling
        $maxStorage = 0
        foreach ($site in $topSites) {
            $storage = [int]($site["Storage"] -replace "[^\d]", "")
            if ($storage -gt $maxStorage) { $maxStorage = $storage }
        }
        
        if ($maxStorage -eq 0) { $maxStorage = 1000 }
        
        $index = 0
        foreach ($site in $topSites) {
            $storage = [int]($site["Storage"] -replace "[^\d]", "")
            $barHeight = ($storage / $maxStorage) * ($chartHeight - 40)
            
            # Determine color based on storage
            $color = switch ($true) {
                ($storage -gt 2000) { "#DC3545" }
                ($storage -gt 1000) { "#FFC107" }
                ($storage -gt 500) { "#17A2B8" }
                default { "#28A745" }
            }

            Write-Host "Assigned color: $color"

            # Remove any non-hex characters like '#'
            $colorHex = $color.TrimStart('#').ToUpper()

            # Ensure the color is a valid 6-character hex code
            if ($colorHex.Length -eq 6 -and $colorHex -match '^[0-9A-F]{6}$') {
                # Convert hex to byte values
                $r = [System.Convert]::ToByte($colorHex.Substring(0, 2), 16)
                $g = [System.Convert]::ToByte($colorHex.Substring(2, 2), 16)
                $b = [System.Convert]::ToByte($colorHex.Substring(4, 2), 16)

                # Create SolidColorBrush using the parsed color
                $brush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromArgb(255, $r, $g, $b))
            } else {
                Write-Error "Invalid color format: $color"
            }
            
            # Create bar
            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $brush
            
            # Position bar
            $x = $startX + ($index * ($barWidth + $spacing))
            $y = $chartHeight - $barHeight - 20
            
            [System.Windows.Controls.Canvas]::SetLeft($bar, $x)
            [System.Windows.Controls.Canvas]::SetTop($bar, $y)
            
            $Canvas.Children.Add($bar)
            
            # Add site label
            $label = New-Object System.Windows.Controls.TextBlock
            $siteTitle = if ($site["Title"]) { $site["Title"] } else { "Site" }
            $label.Text = if ($siteTitle.Length -gt 10) { $siteTitle.Substring(0, 10) + "..." } else { $siteTitle }
            $label.FontSize = 9
            $label.Foreground = "#495057"
            
            [System.Windows.Controls.Canvas]::SetLeft($label, $x)
            [System.Windows.Controls.Canvas]::SetTop($label, $chartHeight - 15)
            
            $Canvas.Children.Add($label)
            
            # Add value label
            $valueLabel = New-Object System.Windows.Controls.TextBlock
            $valueLabel.Text = "$storage MB"
            $valueLabel.FontSize = 8
            $valueLabel.Foreground = "#666"
            
            [System.Windows.Controls.Canvas]::SetLeft($valueLabel, $x + 5)
            [System.Windows.Controls.Canvas]::SetTop($valueLabel, $y - 15)
            
            $Canvas.Children.Add($valueLabel)
            
            $index++
        }
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Draw-StorageDistributionChart"
    }
}





function Load-SiteHealthAnalysis {
    <#
    .SYNOPSIS
    Loads site health analysis
    #>
    param($Controls, $Sites)
    
    try {
        $healthySites = 0
        $warningSites = 0
        $criticalSites = 0
        $healthIssues = @()
        
        foreach ($site in $Sites) {
            $storage = [int]($site["Storage"] -replace "[^\d]", "")
            $hasIssue = $false
            
            # Check for critical storage (>2GB)
            if ($storage -gt 2000) {
                $criticalSites++
                $hasIssue = $true
                $healthIssues += [PSCustomObject]@{
                    Site = $site["Title"]
                    Issue = "Critical storage usage: $storage MB"
                    Recommendation = "Consider archiving old content or increasing storage quota"
                    Icon = "🚨"
                    SeverityColor = "#FFCDD2"
                }
            }
            # Check for high storage (1-2GB)
            elseif ($storage -gt 1000) {
                $warningSites++
                $hasIssue = $true
                $healthIssues += [PSCustomObject]@{
                    Site = $site["Title"]
                    Issue = "High storage usage: $storage MB"
                    Recommendation = "Monitor storage growth and plan for cleanup"
                    Icon = "⚠️"
                    SeverityColor = "#FFE0B2"
                }
            }
            
            # Check for old sites (not modified in 90 days)
            if ($site["LastModified"]) {
                try {
                    $lastMod = [DateTime]::Parse($site["LastModified"])
                    $daysSinceModified = (Get-Date) - $lastMod
                    
                    if ($daysSinceModified.Days -gt 90) {
                        if (-not $hasIssue) { $warningSites++ }
                        $healthIssues += [PSCustomObject]@{
                            Site = $site["Title"]
                            Issue = "Site inactive for $($daysSinceModified.Days) days"
                            Recommendation = "Review if site is still needed or should be archived"
                            Icon = "📅"
                            SeverityColor = "#E1F5FE"
                        }
                        $hasIssue = $true
                    }
                }
                catch { }
            }
            
            if (-not $hasIssue) {
                $healthySites++
            }
        }
        
        # Update health summary
        $Controls.txtHealthySites.Text = $healthySites.ToString()
        $Controls.txtWarningSites.Text = $warningSites.ToString()
        $Controls.txtCriticalSites.Text = $criticalSites.ToString()
        
        # Update health issues list
        $Controls.lstHealthIssues.ItemsSource = $healthIssues
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SiteHealthAnalysis"
    }
}

function Refresh-SitesDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes the deep dive data by fetching latest from SharePoint
    #>
    param($Controls)
    
    try {
        $Controls.txtStatus.Text = "Refreshing data from SharePoint..."
        
        if ($script:SPOConnected) {
            # Call the SPOConnection function to get fresh data
            $sites = Get-DetailedSiteInformation
            
            # Update data manager
            Clear-SharePointData -DataType "Sites"
            foreach ($site in $sites) {
                Add-SharePointSite -SiteData $site
            }
            
            # Reload the window data
            Load-SitesDeepDiveData -Controls $Controls
            
            [System.Windows.MessageBox]::Show(
                "Data refreshed successfully!",
                "Success",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        else {
            [System.Windows.MessageBox]::Show(
                "Not connected to SharePoint. Please connect first.",
                "Warning",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
        }
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-SitesDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-SitesDeepDiveData {
    <#
    .SYNOPSIS
    Exports the deep dive data to CSV
    #>
    param($Controls)
    
    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_Sites_DeepDive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq $true) {
            $sites = Get-SharePointData -DataType "Sites"
            
            $exportData = @()
            foreach ($site in $sites) {
                $exportData += [PSCustomObject]@{
                    Title = $site["Title"]
                    Url = $site["Url"]
                    Owner = $site["Owner"]
                    "Storage (MB)" = $site["Storage"]
                    Template = $site["Template"]
                    LastModified = $site["LastModified"]
                    IsHubSite = $site["IsHubSite"]
                    HasUniquePermissions = $site["HasUniquePermissions"]
                }
            }
            
            $exportData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
            
            [System.Windows.MessageBox]::Show(
                "Data exported successfully to:`n$($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            
            $Controls.txtStatus.Text = "Data exported successfully"
        }
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-SitesDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Apply-SitesFilter {
    <#
    .SYNOPSIS
    Applies filters to the sites data grid
    #>
    param($Controls)
    
    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $templateFilter = $Controls.cboTemplateFilter.SelectedItem.Content
        $storageFilter = $Controls.cboStorageFilter.SelectedItem.Content
        
        $sites = Get-SharePointData -DataType "Sites"
        $filteredSites = @()
        
        foreach ($site in $sites) {
            $includeItem = $true
            
            # Apply search filter
            if ($searchText -and $searchText.Length -gt 0) {
                $title = if ($site["Title"]) { $site["Title"].ToLower() } else { "" }
                $url = if ($site["Url"]) { $site["Url"].ToLower() } else { "" }
                $owner = if ($site["Owner"]) { $site["Owner"].ToLower() } else { "" }
                
                if (-not ($title.Contains($searchText) -or $url.Contains($searchText) -or $owner.Contains($searchText))) {
                    $includeItem = $false
                }
            }
            
            # Apply template filter
            if ($templateFilter -ne "All Templates" -and $includeItem) {
                # Add template filtering logic based on your needs
            }
            
            # Apply storage filter
            if ($storageFilter -ne "All Storage Levels" -and $includeItem) {
                $storage = [int]($site["Storage"] -replace "[^\d]", "")
                
                switch ($storageFilter) {
                    "Low (<500 MB)" { if ($storage -ge 500) { $includeItem = $false } }
                    "Medium (500-1000 MB)" { if ($storage -lt 500 -or $storage -gt 1000) { $includeItem = $false } }
                    "High (1-2 GB)" { if ($storage -lt 1000 -or $storage -gt 2000) { $includeItem = $false } }
                    "Critical (>2 GB)" { if ($storage -le 2000) { $includeItem = $false } }
                }
            }
            
            if ($includeItem) {
                $filteredSites += $site
            }
        }
        
        Load-DetailedSitesGrid -Controls $Controls -Sites $filteredSites
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-SitesFilter"
    }
}

function Get-DetailedSiteInformation {
    <#
    .SYNOPSIS
    Gets detailed site information including real storage data
    #>
    try {
        Write-ActivityLog "Fetching detailed site information for deep dive" -Level "Information"
        
        $detailedSites = @()
        
        if ($script:SPOConnected) {
            $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
            $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
            $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"
            
            # Try admin connection for best results
            $useAdmin = $false
            try {
                Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive -ErrorAction Stop
                $sites = Get-PnPTenantSite -Detailed -ErrorAction Stop
                $useAdmin = $true
                Write-ActivityLog "Using admin connection for detailed site info"
            }
            catch {
                Write-ActivityLog "Admin connection failed, using standard connection"
                $sites = Get-PnPTenantSite -ErrorAction SilentlyContinue
                
                if (-not $sites) {
                    # Fallback to current site
                    $currentWeb = Get-PnPWeb
                    $currentSite = Get-PnPSite -Includes Usage
                    
                    $sites = @([PSCustomObject]@{
                        Title = $currentWeb.Title
                        Url = $currentWeb.Url
                        StorageUsageCurrent = if ($currentSite.Usage) { 
                            [math]::Round($currentSite.Usage.Storage / 1MB, 0) 
                        } else { 0 }
                        Template = $currentWeb.WebTemplate
                        LastContentModifiedDate = $currentWeb.LastItemModifiedDate
                    })
                }
            }
            
            foreach ($site in $sites) {
                # Get real storage value
                $storageValue = 0
                
                if ($site.StorageUsageCurrent -ne $null) {
                    $storageValue = [int]$site.StorageUsageCurrent
                }
                elseif ($site.Url -and -not $useAdmin) {
                    # Try to get storage by connecting to the site
                    try {
                        Connect-PnPOnline -Url $site.Url -ClientId $clientId -Interactive -ErrorAction SilentlyContinue
                        $siteDetail = Get-PnPSite -Includes Usage -ErrorAction SilentlyContinue
                        
                        if ($siteDetail -and $siteDetail.Usage) {
                            $storageValue = [math]::Round($siteDetail.Usage.Storage / 1MB, 0)
                        }
                    }
                    catch {
                        Write-ActivityLog "Could not get storage for: $($site.Url)" -Level "Warning"
                    }
                }
                
                $siteData = @{
                    Title = $site.Title
                    Url = $site.Url
                    Owner = if ($site.Owner) { $site.Owner } else { $site.SiteOwnerEmail }
                    Storage = $storageValue.ToString()  # Real storage value
                    StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
                    Template = if ($site.Template) { $site.Template } else { "N/A" }
                    LastModified = if ($site.LastContentModifiedDate) { 
                        $site.LastContentModifiedDate.ToString("yyyy-MM-dd") 
                    } else { "N/A" }
                    Created = if ($site.Created) { $site.Created.ToString("yyyy-MM-dd") } else { "N/A" }
                    IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
                    HubSiteId = if ($site.HubSiteId) { $site.HubSiteId } else { $null }
                    SharingCapability = if ($site.SharingCapability) { $site.SharingCapability } else { "N/A" }
                    Status = if ($site.Status) { $site.Status } else { "Active" }
                    LockState = if ($site.LockState) { $site.LockState } else { "Unlock" }
                    WebsCount = if ($site.WebsCount) { $site.WebsCount } else { 0 }
                    HasUniquePermissions = $false
                }
                
                # Try to get additional details if we can connect
                if ($site.Url -and -not $useAdmin) {
                    try {
                        Connect-PnPOnline -Url $site.Url -ClientId $clientId -Interactive -ErrorAction SilentlyContinue
                        $web = Get-PnPWeb -ErrorAction SilentlyContinue
                        
                        if ($web) {
                            $siteData["HasUniquePermissions"] = $web.HasUniqueRoleAssignments
                        }
                    }
                    catch {
                        Write-ActivityLog "Could not get additional details for: $($site.Url)" -Level "Warning"
                    }
                }
                
                $detailedSites += $siteData
            }
        }
        
        Write-ActivityLog "Retrieved detailed information for $($detailedSites.Count) sites" -Level "Information"
        return $detailedSites
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-DetailedSiteInformation"
        return @()
    }
}

function Get-SiteHealthMetrics {
    <#
    .SYNOPSIS
    Gets health metrics for a specific site
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )
    
    try {
        Write-ActivityLog "Getting health metrics for site: $SiteUrl" -Level "Information"
        
        $healthMetrics = @{
            SiteUrl = $SiteUrl
            HealthScore = 100  # Start with perfect score
            Issues = @()
            Recommendations = @()
        }
        
        # Connect to the site
        Connect-PnPOnline -Url $SiteUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive
        
        # Check storage usage
        $site = Get-PnPSite
        if ($site.Usage) {
            $storagePercentage = $site.Usage.StoragePercentageUsed
            
            if ($storagePercentage -gt 90) {
                $healthMetrics.HealthScore -= 30
                $healthMetrics.Issues += "Critical: Storage usage above 90%"
                $healthMetrics.Recommendations += "Immediately archive or delete unnecessary content"
            }
            elseif ($storagePercentage -gt 75) {
                $healthMetrics.HealthScore -= 15
                $healthMetrics.Issues += "Warning: Storage usage above 75%"
                $healthMetrics.Recommendations += "Plan for storage cleanup or quota increase"
            }
        }
        
        # Check for orphaned permissions
        $web = Get-PnPWeb
        if ($web.HasUniqueRoleAssignments) {
            $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
            
            # Check for empty groups or invalid principals
            foreach ($assignment in $roleAssignments) {
                try {
                    $member = Get-PnPProperty -ClientObject $assignment -Property Member
                    if (-not $member -or $member.PrincipalType -eq "None") {
                        $healthMetrics.HealthScore -= 10
                        $healthMetrics.Issues += "Orphaned permission assignment detected"
                        $healthMetrics.Recommendations += "Review and clean up permission assignments"
                        break
                    }
                }
                catch { }
            }
        }
        
        # Check for external sharing
        $lists = Get-PnPList
        $externalSharingDetected = $false
        
        foreach ($list in $lists | Where-Object { -not $_.Hidden }) {
            if ($list.HasUniqueRoleAssignments) {
                $healthMetrics.HealthScore -= 5
                $externalSharingDetected = $true
                break
            }
        }
        
        if ($externalSharingDetected) {
            $healthMetrics.Issues += "Multiple lists with unique permissions detected"
            $healthMetrics.Recommendations += "Review list-level permissions for consistency"
        }
        
        # Check last activity
        if ($web.LastItemModifiedDate) {
            $daysSinceModified = (Get-Date) - $web.LastItemModifiedDate
            
            if ($daysSinceModified.Days -gt 180) {
                $healthMetrics.HealthScore -= 20
                $healthMetrics.Issues += "Site inactive for over 180 days"
                $healthMetrics.Recommendations += "Consider archiving or deleting if no longer needed"
            }
            elseif ($daysSinceModified.Days -gt 90) {
                $healthMetrics.HealthScore -= 10
                $healthMetrics.Issues += "Site inactive for over 90 days"
                $healthMetrics.Recommendations += "Review if site is still actively needed"
            }
        }
        
        # Ensure score doesn't go below 0
        if ($healthMetrics.HealthScore -lt 0) { $healthMetrics.HealthScore = 0 }
        
        Write-ActivityLog "Health metrics calculated for site: Score = $($healthMetrics.HealthScore)" -Level "Information"
        return $healthMetrics
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-SiteHealthMetrics"
        return $null
    }
}
