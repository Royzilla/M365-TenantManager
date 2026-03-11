# Graph-Auth.ps1
# Microsoft Graph authentication functions

function Connect-M365 {
    param([Parameter(Mandatory=$true)] $Ui)
    
    if ($script:GraphConnection.Connected) {
        # Disconnect
        try {
            Disconnect-MgGraph | Out-Null
            $script:GraphConnection = @{ Connected = $false; Account = $null; Tenant = $null; TokenExpiry = $null }
            Update-ConnectionStatus -Ui $Ui -Status "Disconnected"
            [System.Windows.MessageBox]::Show("Disconnected from Microsoft 365", "Disconnected", "OK", "Information")
        }
        catch {
            [System.Windows.MessageBox]::Show("Error disconnecting: $_", "Error", "OK", "Error")
        }
        return
    }
    
    try {
        # Connect with interactive login
        $connection = Connect-MgGraph -Scopes @(
            "User.ReadWrite.All"
            "Directory.ReadWrite.All"
            "Organization.Read.All"
            "Group.ReadWrite.All"
            "GroupMember.ReadWrite.All"
            "Directory.Read.All"
        ) -ErrorAction Stop
        
        # Get context info
        $context = Get-MgContext
        $org = Get-MgOrganization
        
        $script:GraphConnection = @{
            Connected = $true
            Account = $context.Account
            Tenant = $org.DisplayName
            TokenExpiry = $context.ExpiresOn
        }
        
        Update-ConnectionStatus -Ui $Ui -Status "Connected" -Account $context.Account -Tenant $org.DisplayName
        Load-DashboardData -Ui $Ui
        
        [System.Windows.MessageBox]::Show("Connected successfully to $($org.DisplayName)!", "Connected", "OK", "Information")
    }
    catch {
        [System.Windows.MessageBox]::Show("Failed to connect: $_", "Connection Error", "OK", "Error")
    }
}

function Test-GraphConnection {
    return $script:GraphConnection.Connected
}

function Assert-GraphConnection {
    if (-not $script:GraphConnection.Connected) {
        [System.Windows.MessageBox]::Show("Please connect to Microsoft 365 first.", "Not Connected", "OK", "Warning")
        return $false
    }
    return $true
}
