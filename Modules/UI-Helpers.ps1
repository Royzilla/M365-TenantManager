# UI-Helpers.ps1
# UI helper functions

function Update-ConnectionStatus {
    param(
        [Parameter(Mandatory=$true)] $Ui,
        [Parameter(Mandatory=$true)] [ValidateSet("Connected", "Disconnected")] $Status,
        [string] $Account = "",
        [string] $Tenant = ""
    )
    
    switch ($Status) {
        "Connected" {
            $Ui.StatusIndicator.Fill = "#FF107C10"
            $Ui.StatusText.Text = "Connected"
            $Ui.StatusAccount.Text = "$Account | $Tenant"
            $Ui.BtnConnect.Content = "Disconnect"
        }
        "Disconnected" {
            $Ui.StatusIndicator.Fill = "#FFE81123"
            $Ui.StatusText.Text = "Disconnected"
            $Ui.StatusAccount.Text = ""
            $Ui.BtnConnect.Content = "Connect to M365"
            
            # Reset stats
            $Ui.StatTotalUsers.Text = "-"
            $Ui.StatActiveUsers.Text = "-"
            $Ui.StatLicenses.Text = "-"
            $Ui.StatGroups.Text = "-"
        }
    }
}

function Set-PageTitle {
    param(
        [Parameter(Mandatory=$true)] $Ui,
        [Parameter(Mandatory=$true)] $Title
    )
    $Ui.PageTitle.Text = $Title
}

function Switch-View {
    param(
        [Parameter(Mandatory=$true)] $Ui,
        [Parameter(Mandatory=$true)] $View
    )
    $Ui.ContentArea.Content = $View
}

function Show-Loading {
    param([string] $Message = "Loading...")
    # Could implement a loading overlay here
}

function Hide-Loading {
    # Hide loading overlay
}
