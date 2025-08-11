function Initialize-HelpTab {
    <#
    .SYNOPSIS
    Initializes the Help tab (minimal logic required as it's mostly static content)
    #>
    try {
        Write-ActivityLog "Initializing Help tab" -Level "Information"
        
        # Help tab is mostly static XAML content, but you could add:
        # - Dynamic version information
        # - Links that open in browser
        # - Copy-to-clipboard functionality for setup instructions
        
        # Example: Add version info dynamically
        Update-HelpTabInfo
        
        Write-ActivityLog "Help tab initialized successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Help tab: $($_.Exception.Message)" -Level "Warning"
        # Help tab failure shouldn't block the application
    }
}

function Update-HelpTabInfo {
    <#
    .SYNOPSIS
    Updates dynamic information in the Help tab
    #>
    try {
        # You could dynamically update version info, last update times, etc.
        # For now, this is just a placeholder for future enhancements
        
        Write-ActivityLog "Help tab information updated" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error updating help tab info: $($_.Exception.Message)" -Level "Warning"
    }
}