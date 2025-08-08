function Show-MainWindow {
    try {
        $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="SharePoint Permissions Tool" Height="650" Width="950"
        WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="15">
            <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="SharePoint Online Permissions Tool" 
                          FontSize="18" FontWeight="Bold" Foreground="White" 
                          HorizontalAlignment="Center"/>
                <TextBlock Text="App Registration Authentication" 
                          FontSize="12" Foreground="White" 
                          HorizontalAlignment="Center" Margin="0,2,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="10">
            <!-- Connection Tab -->
            <TabItem Header="Connection">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Tenant URL Input -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,15">
                        <TextBlock Text="SharePoint Tenant URL:" Margin="0,0,0,5" FontWeight="SemiBold"/>
                        <TextBox Name="txtTenantUrl" Height="30" FontSize="12"
                                Text="https://yourtenant.sharepoint.com"/>
                        <TextBlock Text="Example: https://contoso.sharepoint.com" 
                                  FontSize="10" Foreground="Gray" Margin="0,2,0,0"/>
                    </StackPanel>
                    
                    <!-- Client ID Input -->
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="App Registration Client ID:" Margin="0,0,0,5" FontWeight="SemiBold"/>
                        <TextBox Name="txtClientId" Height="30" FontSize="12" Text=""/>
                        <TextBlock Text="GUID format: 12345678-1234-1234-1234-123456789012" 
                                  FontSize="10" Foreground="Gray" Margin="0,2,0,0"/>
                    </StackPanel>
                    
                    <!-- Info Panel -->
                    <Border Grid.Row="2" Background="#F0F8FF" BorderBrush="#0078D4" 
                           BorderThickness="1" Padding="10" Margin="0,0,0,15">
                        <StackPanel>
                            <TextBlock Text="App Registration Requirements:" FontWeight="SemiBold" 
                                      Foreground="#0078D4" Margin="0,0,0,5"/>
                            <TextBlock Text="• Redirect URI: http://localhost" FontSize="11" Margin="0,0,0,2"/>
                            <TextBlock Text="• API Permissions: Sites.FullControl.All, User.Read.All" FontSize="11" Margin="0,0,0,2"/>
                            <TextBlock Text="• Grant admin consent for the permissions" FontSize="11" Margin="0,0,0,2"/>
                            <TextBlock Text="• Set as public client (Allow public client flows = Yes)" FontSize="11"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Buttons -->
                    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,15">
                        <Button Name="btnConnectSPO" Content="Connect to SharePoint" 
                               Width="200" Height="35" Background="#0078D4" Foreground="White"
                               FontWeight="SemiBold" Margin="0,0,10,0"/>
                        <Button Name="btnDemoMode" Content="Demo Mode" 
                               Width="120" Height="35" Margin="0,0,0,0"/>
                    </StackPanel>
                    
                    <!-- Results -->
                    <ScrollViewer Grid.Row="4">
                        <TextBox Name="txtSPOResults" IsReadOnly="True" 
                                TextWrapping="Wrap" AcceptsReturn="True"
                                Background="#F5F5F5" FontFamily="Consolas" FontSize="11"
                                Text="Enter your SharePoint tenant URL and App Registration Client ID, then click Connect to SharePoint."/>
                    </ScrollViewer>
                </Grid>
            </TabItem>
            
            <!-- SharePoint Operations Tab -->
            <TabItem Header="SharePoint Operations">
                <Grid Margin="20">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Connection Status -->
                    <Border Grid.Row="0" Background="#F0F0F0" Padding="10" Margin="0,0,0,15">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Connection Status: " FontWeight="SemiBold"/>
                            <TextBlock Name="txtStatus" Text="Not connected to SharePoint" 
                                      Foreground="Red"/>
                        </StackPanel>
                    </Border>
                    
                    <!-- Site URL Input -->
                    <StackPanel Grid.Row="1" Margin="0,0,0,15">
                        <TextBlock Text="Site URL for Permissions Analysis (Optional):" Margin="0,0,0,5" FontWeight="SemiBold"/>
                        <TextBox Name="txtSiteUrl" Height="30" FontSize="12"/>
                        <TextBlock Text="Leave empty for tenant-level operations, or enter specific site URL" 
                                  FontSize="10" Foreground="Gray" Margin="0,2,0,0"/>
                    </StackPanel>
                    
                    <!-- Action Buttons -->
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,15">
                        <Button Name="btnGetSites" Content="Get All Sites" Width="120" Height="30" 
                               Margin="0,0,10,0" IsEnabled="False"/>
                        <Button Name="btnGetPermissions" Content="Analyze Permissions" Width="140" Height="30" 
                               Margin="0,0,10,0" IsEnabled="False"/>
                        <Button Name="btnGenerateReport" Content="Generate Report" Width="120" Height="30" 
                               IsEnabled="False"/>
                    </StackPanel>
                    
                    <!-- Help Text -->
                    <TextBlock Grid.Row="3" Text="Operations available after connecting to SharePoint:" 
                              FontStyle="Italic" Foreground="Gray" Margin="0,0,0,10"/>
                    
                    <!-- Operations Results -->
                    <ScrollViewer Grid.Row="4">
                        <TextBox Name="txtOperationsResults" IsReadOnly="True" 
                                TextWrapping="Wrap" AcceptsReturn="True"
                                Background="#F5F5F5" FontFamily="Consolas" FontSize="11"
                                Text="Connect to SharePoint to begin operations..."/>
                    </ScrollViewer>
                </Grid>
            </TabItem>
            
            <!-- Help Tab -->
            <TabItem Header="Help">
                <ScrollViewer Margin="20">
                    <StackPanel>
                        <TextBlock Text="SharePoint Permissions Tool - Help" FontSize="16" FontWeight="Bold" 
                                  Margin="0,0,0,15"/>
                        
                        <TextBlock Text="App Registration Setup:" FontSize="14" FontWeight="SemiBold" 
                                  Margin="0,0,0,10"/>
                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                            1. Go to Azure Portal > App registrations > New registration<LineBreak/>
                            2. Name: 'SharePoint Permissions Tool'<LineBreak/>
                            3. Supported account types: Single tenant<LineBreak/>
                            4. Redirect URI: Web > http://localhost<LineBreak/>
                            5. Register the application
                        </TextBlock>
                        
                        <TextBlock Text="Configure Permissions:" FontSize="14" FontWeight="SemiBold" 
                                  Margin="0,0,0,10"/>
                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                            1. Go to API permissions<LineBreak/>
                            2. Add permission > Microsoft Graph > Delegated permissions<LineBreak/>
                            3. Add: Sites.FullControl.All, User.Read.All, Group.Read.All<LineBreak/>
                            4. Add permission > SharePoint > Delegated permissions<LineBreak/>
                            5. Add: AllSites.FullControl<LineBreak/>
                            6. Grant admin consent for your organization
                        </TextBlock>
                        
                        <TextBlock Text="Enable Public Client:" FontSize="14" FontWeight="SemiBold" 
                                  Margin="0,0,0,10"/>
                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                            1. Go to Authentication<LineBreak/>
                            2. Scroll down to 'Allow public client flows'<LineBreak/>
                            3. Set to 'Yes'<LineBreak/>
                            4. Save
                        </TextBlock>
                        
                        <TextBlock Text="Usage:" FontSize="14" FontWeight="SemiBold" 
                                  Margin="0,0,0,10"/>
                        <TextBlock TextWrapping="Wrap" Margin="0,0,0,15">
                            1. Copy the Application (client) ID from your app registration<LineBreak/>
                            2. Enter your SharePoint tenant URL and Client ID<LineBreak/>
                            3. Click 'Connect to SharePoint'<LineBreak/>
                            4. Complete authentication in the popup window<LineBreak/>
                            5. Use the SharePoint Operations tab to analyze permissions
                        </TextBlock>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
        
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
        $window = [System.Windows.Markup.XamlReader]::Load($reader)
        
        # Get controls and store as script variables
        $script:txtTenantUrl = $window.FindName("txtTenantUrl")
        $script:txtClientId = $window.FindName("txtClientId")
        $script:btnConnectSPO = $window.FindName("btnConnectSPO")
        $script:btnDemoMode = $window.FindName("btnDemoMode")
        $script:txtSPOResults = $window.FindName("txtSPOResults")
        $script:txtSiteUrl = $window.FindName("txtSiteUrl")
        $script:btnGetSites = $window.FindName("btnGetSites")
        $script:btnGetPermissions = $window.FindName("btnGetPermissions")
        $script:btnGenerateReport = $window.FindName("btnGenerateReport")
        $script:txtStatus = $window.FindName("txtStatus")
        $script:txtOperationsResults = $window.FindName("txtOperationsResults")
        
        # Load saved settings
        $savedTenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
        if ($savedTenantUrl) {
            $script:txtTenantUrl.Text = $savedTenantUrl
        }
        
        $savedClientId = Get-AppSetting -SettingName "SharePoint.ClientId"
        if ($savedClientId) {
            $script:txtClientId.Text = $savedClientId
        }
        
        # Event handlers
        $script:btnConnectSPO.Add_Click({ Connect-SPO })
        
        $script:btnDemoMode.Add_Click({
            Set-AppSetting -SettingName "DemoMode" -Value $true
            $script:SPOConnected = $true
            $script:txtSPOResults.Text = "Demo Mode Enabled - SharePoint Connected (Simulated)`nApp Registration: demo-client-id`nAuthenticated User: demo@contoso.com"
            $script:txtStatus.Text = "Connected (Demo Mode)"
            $script:txtStatus.Foreground = "Green"
            $script:btnGetSites.IsEnabled = $true
            $script:btnGetPermissions.IsEnabled = $true
            $script:btnGenerateReport.IsEnabled = $true
        })
        
        $script:btnGetSites.Add_Click({ 
            if (Get-AppSetting -SettingName "DemoMode") {
                $demoSitesText = @"
Demo SharePoint Sites:

Sites Found: 5

Title: Team Site
URL: https://demo.sharepoint.com/sites/teamsite
Owner: admin@demo.com
Storage Used: 245 MB
Template: STS#3
---

Title: Project Alpha
URL: https://demo.sharepoint.com/sites/project-alpha
Owner: project.manager@demo.com
Storage Used: 1024 MB
Template: PROJECTSITE#0
---

Title: HR Documents
URL: https://demo.sharepoint.com/sites/hr-documents
Owner: hr.admin@demo.com
Storage Used: 512 MB
Template: STS#3
---

Title: Company Intranet
URL: https://demo.sharepoint.com
Owner: admin@demo.com
Storage Used: 2048 MB
Template: SITEPAGEPUBLISHING#0
---

Title: Sales Team
URL: https://demo.sharepoint.com/sites/sales
Owner: sales.manager@demo.com
Storage Used: 768 MB
Template: STS#3
---
"@
                $script:txtOperationsResults.Text = $demoSitesText
            } else {
                Get-SharePointSites
            }
        })
        
        $script:btnGetPermissions.Add_Click({ 
            if (Get-AppSetting -SettingName "DemoMode") {
                $siteUrl = if ($script:txtSiteUrl.Text.Trim()) { $script:txtSiteUrl.Text.Trim() } else { "https://demo.sharepoint.com/sites/teamsite" }
                $demoPermissionsText = @"
Demo Permissions Analysis for: $siteUrl
Site Title: Team Collaboration Site
Site Description: Main team collaboration workspace

Users: 15
Groups: 6

Site Users (excluding system accounts):
- John Doe (john.doe@demo.com)
- Jane Smith (jane.smith@demo.com)
- Mike Johnson (mike.johnson@demo.com)
- Sarah Wilson (sarah.wilson@demo.com)
- Alex Brown (alex.brown@demo.com)
- Lisa Davis (lisa.davis@demo.com)
- Tom Miller (tom.miller@demo.com)
- Emma Taylor (emma.taylor@demo.com)
- David Anderson (david.anderson@demo.com)
- Maria Garcia (maria.garcia@demo.com)

SharePoint Groups:
- Team Site Owners (3 members)
- Team Site Members (8 members)
- Team Site Visitors (4 members)
- Project Alpha Contributors (6 members)
- Document Reviewers (5 members)
- External Partners (2 members)

... showing first 10 entries of each type
"@
                $script:txtOperationsResults.Text = $demoPermissionsText
            } else {
                Get-SharePointPermissions
            }
        })
        
        $script:btnGenerateReport.Add_Click({
            if (Get-AppSetting -SettingName "DemoMode") {
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $demoReportText = @"
Demo Report Generation Complete!

Report Details:
- Report Name: SharePoint_Permissions_Report_$timestamp.csv
- Location: ./Reports/SharePoint_Permissions_Report_$timestamp.csv
- Format: CSV (Excel compatible)

Report Contains:
✓ Site Collection Information (5 sites)
✓ User Permissions Summary (47 unique users)
✓ Group Memberships (23 SharePoint groups)
✓ Permission Inheritance Analysis
✓ External User Access Report
✓ Orphaned Permissions Detection

Additional Files Created:
- SharePoint_Sites_Summary_$timestamp.csv
- User_Access_Matrix_$timestamp.xlsx
- Permissions_Audit_Log_$timestamp.txt

Report Generation Time: 2.3 seconds
Total Records Processed: 1,247

The reports are ready for review and can be opened in Excel or imported into other systems.
"@
                $script:txtOperationsResults.Text = $demoReportText
            } else {
                # Real report generation would go here
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $reportPath = "./Reports/PermissionsReport_$timestamp.csv"
                $script:txtOperationsResults.Text = "Generating permissions report...`n`nReport will be saved to: $reportPath`n`nThis feature is under development. In Demo Mode, you can see a sample of what the report would contain."
            }
        })
        
        # Show window
        $window.ShowDialog() | Out-Null
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-MainWindow"
        throw
    }
}