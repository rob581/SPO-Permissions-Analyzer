function New-UIManager {
    param($ResultsBox, $ConnectButton, $Controls)

    $uiManager = [PSCustomObject]@{
        ResultsBox = $ResultsBox
        ConnectButton = $ConnectButton
        Controls = $Controls
        Dispatcher = $ResultsBox.Dispatcher
        IsConnected = $false
    }

    # Add methods
    Add-Member -InputObject $uiManager -MemberType ScriptMethod -Name "UpdateStatus" -Value {
        param($Message)
        if ($null -ne $this.Dispatcher) {
            $this.Dispatcher.Invoke([Action]{
                $this.ResultsBox.Text = $Message
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
        }
    }

    Add-Member -InputObject $uiManager -MemberType ScriptMethod -Name "EnableControls" -Value {
        $this.IsConnected = $true
        if ($null -ne $this.Dispatcher) {
            $this.Dispatcher.Invoke([Action]{
                $this.ConnectButton.IsEnabled = $false
                foreach ($control in $this.Controls.Values) {
                    if ($control) { $control.IsEnabled = $true }
                }
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
        }
    }

    Add-Member -InputObject $uiManager -MemberType ScriptMethod -Name "DisableControls" -Value {
        $this.IsConnected = $false
        if ($null -ne $this.Dispatcher) {
            $this.Dispatcher.Invoke([Action]{
                $this.ConnectButton.IsEnabled = $true
                foreach ($control in $this.Controls.Values) {
                    if ($control) { $control.IsEnabled = $false }
                }
            }, [System.Windows.Threading.DispatcherPriority]::Normal)
        }
    }

    return $uiManager
}
