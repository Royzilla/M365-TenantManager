# User-Management.ps1
# User management functions and views

function Load-DashboardData {
    param([Parameter(Mandatory=$true)] $Ui)
    
    try {
        # Get user counts
        $totalUsers = (Get-MgUser -All).Count
        $activeUsers = (Get-MgUser -Filter "accountEnabled eq true" -All).Count
        
        # Get license count
        $skus = Get-MgSubscribedSku
        $totalLicenses = ($skus | Measure-Object -Property ConsumedUnits -Sum).Sum
        
        # Get group count
        $groups = (Get-MgGroup -All).Count
        
        # Update UI
        $Ui.StatTotalUsers.Text = $totalUsers.ToString()
        $Ui.StatActiveUsers.Text = $activeUsers.ToString()
        $Ui.StatLicenses.Text = $totalLicenses.ToString()
        $Ui.StatGroups.Text = $groups.ToString()
    }
    catch {
        Write-Warning "Failed to load dashboard data: $_"
    }
}

function Show-DashboardView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "Dashboard"
    
    # Reload data if connected
    if ($script:GraphConnection.Connected) {
        Load-DashboardData -Ui $Ui
    }
}

function Show-UserManagementView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "User Management"
    
    if (-not (Assert-GraphConnection)) { return }
    
    # Create user management view
    [xml]$userXaml = @"
&lt;Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"&gt;
    &lt;Grid.RowDefinitions&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="*"/&gt;
    &lt;/Grid.RowDefinitions&gt;
    
    &lt;WrapPanel Grid.Row="0" Margin="0,0,0,15"&gt;
        &lt;Button x:Name="btnBulkImport" Style="{StaticResource ModernButton}" Content="📊 Bulk Import from Excel" Margin="0,0,10,0"/&gt;
        &lt;Button x:Name="btnNewUser" Style="{StaticResource ModernButton}" Content="➕ Create New User" Margin="0,0,10,0"/&gt;
        &lt;Button x:Name="btnExportUsers" Style="{StaticResource ModernButton}" Content="📤 Export to Excel"/&gt;
    &lt;/WrapPanel&gt;
    
    &lt;Border Grid.Row="1" Style="{StaticResource CardBorder}" Padding="10"&gt;
        &lt;StackPanel Orientation="Horizontal"&gt;
            &lt;TextBox x:Name="txtSearch" Width="300" Style="{StaticResource ModernTextBox}" Margin="0,0,10,0"
                    Text="Search users..."/&gt;
            &lt;ComboBox x:Name="cmbFilter" Width="150" Style="{StaticResource ModernComboBox}" Margin="0,0,10,0"&gt;
                &lt;ComboBoxItem Content="All Users" IsSelected="True"/&gt;
                &lt;ComboBoxItem Content="Active Users"/&gt;
                &lt;ComboBoxItem Content="Guest Users"/&gt;
            &lt;/ComboBox&gt;
            &lt;Button x:Name="btnSearch" Style="{StaticResource ModernButton}" Content="🔍 Search"/&gt;
        &lt;/StackPanel&gt;
    &lt;/Border&gt;
    
    &lt;Border Grid.Row="2" Style="{StaticResource CardBorder}"&gt;
        &lt;DataGrid x:Name="dgUsers" AutoGenerateColumns="False" IsReadOnly="True"
                 GridLinesVisibility="Horizontal" Background="Transparent"
                 RowBackground="&#35;FF2D2D30" AlternatingRowBackground="&#35;FF252526"
                 Foreground="White" BorderThickness="0"
                 HeadersVisibility="Column" CanUserAddRows="False"&gt;
            &lt;DataGrid.Columns&gt;
                &lt;DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="*"/&gt;
                &lt;DataGridTextColumn Header="Email" Binding="{Binding UserPrincipalName}" Width="*"/&gt;
                &lt;DataGridTextColumn Header="Department" Binding="{Binding Department}" Width="150"/&gt;
                &lt;DataGridTextColumn Header="Job Title" Binding="{Binding JobTitle}" Width="150"/&gt;
                &lt;DataGridCheckBoxColumn Header="Active" Binding="{Binding AccountEnabled}" Width="70"/&gt;
            &lt;/DataGrid.Columns&gt;
        &lt;/DataGrid&gt;
    &lt;/Border&gt;
&lt;/Grid&gt;
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $userXaml
    $view = [Windows.Markup.XamlReader]::Load($reader)
    
    # Wire up buttons
    $view.FindName("btnBulkImport").Add_Click({ Show-BulkImportView -Ui $Ui })
    $view.FindName("btnNewUser").Add_Click({ Show-CreateUserView -Ui $Ui })
    $view.FindName("btnExportUsers").Add_Click({ Export-UsersToExcel })
    $view.FindName("btnSearch").Add_Click({ 
        Search-Users -DataGrid $view.FindName("dgUsers") -SearchText $view.FindName("txtSearch").Text 
    })
    
    # Load initial user list
    Load-Users -DataGrid $view.FindName("dgUsers")
    
    Switch-View -Ui $Ui -View $view
}

function Show-BulkImportView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "Bulk Import Users"
    
    if (-not (Assert-GraphConnection)) { return }
    
    [xml]$bulkXaml = @"
&lt;Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"&gt;
    &lt;Grid.RowDefinitions&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="*"/&gt;
        &lt;RowDefinition Height="Auto"/&gt;
    &lt;/Grid.RowDefinitions&gt;
    
    &lt;Border Grid.Row="0" Style="{StaticResource CardBorder}"&gt;
        &lt;StackPanel&gt;
            &lt;TextBlock Text="1. Select Excel File" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/&gt;
            &lt;StackPanel Orientation="Horizontal"&gt;
                &lt;TextBox x:Name="txtFilePath" Width="400" Style="{StaticResource ModernTextBox}" IsReadOnly="True"/&gt;
                &lt;Button x:Name="btnBrowse" Style="{StaticResource ModernButton}" Content="Browse..." Margin="10,0,0,0"/&gt;
            &lt;/StackPanel&gt;
        &lt;/StackPanel&gt;
    &lt;/Border&gt;
    
    &lt;Border Grid.Row="1" Style="{StaticResource CardBorder}"&gt;
        &lt;StackPanel&gt;
            &lt;TextBlock Text="2. Import Options" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/&gt;
            
            &lt;CheckBox x:Name="chkForcePasswordChange" Style="{StaticResource ModernCheckBox}" 
                     Content="Force password change at next sign-in" IsChecked="True"/&gt;
            
            &lt;CheckBox x:Name="chkAssignLicense" Style="{StaticResource ModernCheckBox}" 
                     Content="Assign Microsoft 365 license"&gt;
                &lt;StackPanel Margin="20,5,0,0" IsEnabled="{Binding ElementName=chkAssignLicense, Path=IsChecked}"&gt;
                    &lt;ComboBox x:Name="cmbLicense" Width="300" Style="{StaticResource ModernComboBox}"&gt;
                        &lt;ComboBoxItem Content="Microsoft 365 Business Basic"/&gt;
                        &lt;ComboBoxItem Content="Microsoft 365 Business Standard"/&gt;
                        &lt;ComboBoxItem Content="Microsoft 365 Business Premium"/&gt;
                        &lt;ComboBoxItem Content="Office 365 E3"/&gt;
                        &lt;ComboBoxItem Content="Office 365 E5"/&gt;
                    &lt;/ComboBox&gt;
                &lt;/StackPanel&gt;
            &lt;/CheckBox&gt;
            
            &lt;CheckBox x:Name="chkDryRun" Style="{StaticResource ModernCheckBox}" 
                     Content="Dry Run (preview only - don't create users)" IsChecked="True"
                     Foreground="&#35;FFFFA500"/&gt;
        &lt;/StackPanel&gt;
    &lt;/Border&gt;
    
    &lt;Border Grid.Row="2" Style="{StaticResource CardBorder}"&gt;
        &lt;Grid&gt;
            &lt;Grid.RowDefinitions&gt;
                &lt;RowDefinition Height="Auto"/&gt;
                &lt;RowDefinition Height="*"/&gt;
            &lt;/Grid.RowDefinitions&gt;
            
            &lt;TextBlock Grid.Row="0" Text="3. Preview &amp; Import" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/&gt;
            
            &lt;DataGrid x:Name="dgPreview" Grid.Row="1" AutoGenerateColumns="True" IsReadOnly="True"
                     GridLinesVisibility="Horizontal" Background="Transparent"
                     RowBackground="&#35;FF2D2D30" AlternatingRowBackground="&#35;FF252526"
                     Foreground="White" BorderThickness="0" HeadersVisibility="Column"/&gt;
        &lt;/Grid&gt;
    &lt;/Border&gt;
    
    &lt;StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0"&gt;
        &lt;Button x:Name="btnPreview" Style="{StaticResource ModernButton}" Content="👁️ Preview Data" Margin="0,0,10,0"/&gt;
        &lt;Button x:Name="btnImport" Style="{StaticResource ModernButton}" Content="🚀 Import Users" Background="&#35;FF107C10"/&gt;
    &lt;/StackPanel&gt;
&lt;/Grid&gt;
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $bulkXaml
    $view = [Windows.Markup.XamlReader]::Load($reader)
    
    # File path variable
    $script:importFilePath = $null
    $script:importData = $null
    
    # Wire up events
    $view.FindName("btnBrowse").Add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv"
        $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        
        if ($dialog.ShowDialog() -eq "OK") {
            $view.FindName("txtFilePath").Text = $dialog.FileName
            $script:importFilePath = $dialog.FileName
        }
    })
    
    $view.FindName("btnPreview").Add_Click({
        if (-not $script:importFilePath) {
            [System.Windows.MessageBox]::Show("Please select a file first.", "No File", "OK", "Warning")
            return
        }
        
        try {
            $script:importData = Import-Excel -Path $script:importFilePath
            $view.FindName("dgPreview").ItemsSource = $script:importData
            [System.Windows.MessageBox]::Show("Loaded $($script:importData.Count) users.", "Preview Loaded", "OK", "Information")
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to load file: $_", "Error", "OK", "Error")
        }
    })
    
    $view.FindName("btnImport").Add_Click({
        if (-not $script:importData) {
            [System.Windows.MessageBox]::Show("Please preview the data first.", "No Data", "OK", "Warning")
            return
        }
        
        $dryRun = $view.FindName("chkDryRun").IsChecked
        $forcePasswordChange = $view.FindName("chkForcePasswordChange").IsChecked
        $assignLicense = $view.FindName("chkAssignLicense").IsChecked
        $selectedLicense = $view.FindName("cmbLicense").SelectedItem
        
        # Confirm
        $mode = if ($dryRun) { "DRY RUN" } else { "LIVE IMPORT" }
        $confirm = [System.Windows.MessageBox]::Show(
            "About to import $($script:importData.Count) user(s).`n`nMode: $mode`n`nContinue?",
            "Confirm Import", "YesNo", if ($dryRun) { "Information" } else { "Warning" })
        
        if ($confirm -ne "Yes") { return }
        
        # Import users
        Import-UsersFromData -Data $script:importData -DryRun $dryRun -ForcePasswordChange $forcePasswordChange `
                            -AssignLicense $assignLicense -LicenseSku $selectedLicense.Content
    })
    
    Switch-View -Ui $Ui -View $view
}

function Import-UsersFromData {
    param(
        $Data,
        [bool]$DryRun,
        [bool]$ForcePasswordChange,
        [bool]$AssignLicense,
        [string]$LicenseSku
    )
    
    $successCount = 0
    $failCount = 0
    $results = @()
    
    foreach ($user in $Data) {
        try {
            if ($DryRun) {
                $results += "[DRY RUN] Would create: $($user.DisplayName)"
                $successCount++
                continue
            }
            
            # Create password profile
            $passwordProfile = @{
                Password = $user.Password
                ForceChangePasswordNextSignIn = $ForcePasswordChange
            }
            
            # Build user parameters
            $userParams = @{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                MailNickname = $user.MailNickname
                AccountEnabled = $true
                PasswordProfile = $passwordProfile
            }
            
            # Optional fields
            if ($user.GivenName) { $userParams.GivenName = $user.GivenName }
            if ($user.Surname) { $userParams.Surname = $user.Surname }
            if ($user.JobTitle) { $userParams.JobTitle = $user.JobTitle }
            if ($user.Department) { $userParams.Department = $user.Department }
            if ($user.OfficeLocation) { $userParams.OfficeLocation = $user.OfficeLocation }
            if ($user.MobilePhone) { $userParams.MobilePhone = $user.MobilePhone }
            if ($user.BusinessPhones) { $userParams.BusinessPhones = @($user.BusinessPhones) }
            if ($user.StreetAddress) { $userParams.StreetAddress = $user.StreetAddress }
            if ($user.City) { $userParams.City = $user.City }
            if ($user.State) { $userParams.State = $user.State }
            if ($user.PostalCode) { $userParams.PostalCode = $user.PostalCode }
            if ($user.Country) { $userParams.Country = $user.Country }
            if ($user.UsageLocation) { $userParams.UsageLocation = $user.UsageLocation }
            
            # Create user
            $newUser = New-MgUser @userParams
            
            # Assign license if requested
            if ($AssignLicense -and $LicenseSku) {
                $skuPartNumber = Convert-LicenseNameToSku -LicenseName $LicenseSku
                if ($skuPartNumber) {
                    $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $skuPartNumber }
                    if ($sku) {
                        $license = @{ AddLicenses = @(@{ SkuId = $sku.SkuId }); RemoveLicenses = @() }
                        Set-MgUserLicense -UserId $newUser.Id -BodyParameter $license | Out-Null
                    }
                }
            }
            
            # Add to groups
            if ($user.Groups) {
                $groupList = $user.Groups -split ";"
                foreach ($groupName in $groupList) {
                    $group = Get-MgGroup -Filter "displayName eq '$($groupName.Trim())'" -ErrorAction SilentlyContinue
                    if ($group) {
                        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id -ErrorAction SilentlyContinue | Out-Null
                    }
                }
            }
            
            $results += "✓ Created: $($user.DisplayName)"
            $successCount++
        }
        catch {
            $results += "✗ Failed: $($user.DisplayName) - $($_.Exception.Message)"
            $failCount++
        }
    }
    
    # Show results
    $resultMessage = ($results -join "`n") + "`n`nImport complete!`nSuccessful: $successCount`nFailed: $failCount"
    [System.Windows.MessageBox]::Show($resultMessage, "Import Complete", "OK", 
        $(if ($failCount -eq 0) { "Information" } else { "Warning" }))
}

function Load-Users {
    param($DataGrid)
    
    try {
        $users = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, Department, JobTitle, AccountEnabled
        $DataGrid.ItemsSource = $users
    }
    catch {
        Write-Warning "Failed to load users: $_"
    }
}

function Search-Users {
    param($DataGrid, $SearchText)
    
    try {
        $filter = "startsWith(displayName,'$SearchText') or startsWith(userPrincipalName,'$SearchText')"
        $users = Get-MgUser -Filter $filter | Select-Object DisplayName, UserPrincipalName, Department, JobTitle, AccountEnabled
        $DataGrid.ItemsSource = $users
    }
    catch {
        Write-Warning "Search failed: $_"
    }
}

function Export-UsersToExcel {
    if (-not (Assert-GraphConnection)) { return }
    
    try {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
        $saveDialog.FileName = "M365-Users-$(Get-Date -Format 'yyyyMMdd').xlsx"
        $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            $users = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName, MailNickname, 
                GivenName, Surname, JobTitle, Department, OfficeLocation, MobilePhone, 
                BusinessPhones, AccountEnabled, CreatedDateTime
            
            $users | Export-Excel -Path $saveDialog.FileName -AutoSize -TableName "Users"
            [System.Windows.MessageBox]::Show("Exported to $($saveDialog.FileName)", "Export Complete", "OK", "Information")
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Export failed: $_", "Error", "OK", "Error")
    }
}

function Show-CreateUserView {
    param([Parameter(Mandatory=$true)] $Ui)
    
    [System.Windows.MessageBox]::Show("Single user creation dialog coming in next update!`n`nFor now, please use Bulk Import with a single-row Excel file.", 
        "Feature Coming Soon", "OK", "Information")
}

function Convert-LicenseNameToSku {
    param([string]$LicenseName)
    
    switch ($LicenseName) {
        "Microsoft 365 Business Basic" { return "O365_BUSINESS_ESSENTIALS" }
        "Microsoft 365 Business Standard" { return "O365_BUSINESS_PREMIUM" }
        "Microsoft 365 Business Premium" { return "SPB" }
        "Office 365 E3" { return "ENTERPRISEPACK" }
        "Office 365 E5" { return "ENTERPRISEPREMIUM" }
        default { return $null }
    }
}
