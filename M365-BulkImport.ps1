<#
.SYNOPSIS
    M365 Bulk User Import - Import users from Excel to Azure AD with validation

.DESCRIPTION
    PowerShell GUI tool for bulk importing users to Microsoft 365 with
    comprehensive validation, duplicate checking, and detailed reporting.

.EXAMPLE
    .\M365-BulkImport.ps1
#>

#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, ImportExcel

# Error handling
$ErrorActionPreference = "Stop"

# Add required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region XAML UI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Bulk User Import" Height="800" Width="1000"
        WindowStartupLocation="CenterScreen"
        Background="#FF1E1E1E"
        Foreground="White"
        FontFamily="Segoe UI">
    
    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#FF0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,10"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF106EBE"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#FF6C6C6C"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="#FF252526"/>
            <Setter Property="BorderBrush" Value="#FF3E3E42"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="20"/>
            <Setter Property="Margin" Value="0,0,0,15"/>
        </Style>
        
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Background" Value="#FF3C3C3C"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#FF555555"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        
        <Style x:Key="ModernComboBox" TargetType="ComboBox">
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,8"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <StackPanel Grid.Row="0" Margin="0,0,0,20">
            <TextBlock Text="M365 Bulk User Import" FontSize="28" FontWeight="Bold" Foreground="#FF0078D4"/>
            <TextBlock Text="Import users from Excel to Microsoft 365" FontSize="14" Foreground="#FFAAAAAA" Margin="0,5,0,0"/>
        </StackPanel>
        
        <!-- Connection Status -->
        <Border Grid.Row="1" Style="{StaticResource CardBorder}">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <StackPanel Grid.Column="0">
                    <TextBlock Text="Connection Status" FontSize="12" Foreground="#FF858585" Margin="0,0,0,5"/>
                    <StackPanel Orientation="Horizontal">
                        <Ellipse x:Name="statusIndicator" Width="12" Height="12" Fill="#FFE81123" Margin="0,0,10,0"/>
                        <TextBlock x:Name="statusText" Text="Not Connected" FontSize="14" Foreground="White"/>
                    </StackPanel>
                    <TextBlock x:Name="statusAccount" Text="" FontSize="11" Foreground="#FF858585" Margin="22,3,0,0"/>
                </StackPanel>
                
                <Button x:Name="btnConnect" Grid.Column="1" Style="{StaticResource ModernButton}" 
                        Content="Connect to M365" Width="150"/>
            </Grid>
        </Border>
        
        <!-- File Selection -->
        <Border Grid.Row="2" Style="{StaticResource CardBorder}">
            <StackPanel>
                <TextBlock Text="1. Select Excel File" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <StackPanel Orientation="Horizontal">
                    <TextBox x:Name="txtFilePath" Width="550" Style="{StaticResource ModernTextBox}" IsReadOnly="True"/>
                    <Button x:Name="btnBrowse" Style="{StaticResource ModernButton}" Content="Browse..." Margin="10,0,0,0" Width="100"/>
                    <Button x:Name="btnDownloadTemplate" Style="{StaticResource ModernButton}" Content="Download Template" Margin="10,0,0,0" Width="130" Background="#FF6C6C6C"/>
                </StackPanel>
            </StackPanel>
        </Border>
        
        <!-- Options -->
        <Border Grid.Row="3" Style="{StaticResource CardBorder}">
            <StackPanel>
                <TextBlock Text="2. Import Options" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>
                
                <CheckBox x:Name="chkValidateData" Style="{StaticResource ModernCheckBox}" 
                         Content="Validate data before import (recommended)" IsChecked="True"/>
                
                <CheckBox x:Name="chkCheckExisting" Style="{StaticResource ModernCheckBox}" 
                         Content="Check if users already exist" IsChecked="True"/>
                
                <CheckBox x:Name="chkForcePasswordChange" Style="{StaticResource ModernCheckBox}" 
                         Content="Force password change at next sign-in" IsChecked="True"/>
                
                <CheckBox x:Name="chkAssignLicense" Style="{StaticResource ModernCheckBox}" 
                         Content="Assign Microsoft 365 license">
                    <StackPanel Margin="20,5,0,0" IsEnabled="{Binding ElementName=chkAssignLicense, Path=IsChecked}">
                        <ComboBox x:Name="cmbLicense" Width="300" Style="{StaticResource ModernComboBox}">
                            <ComboBoxItem Content="Microsoft 365 Business Basic"/>
                            <ComboBoxItem Content="Microsoft 365 Business Standard"/>
                            <ComboBoxItem Content="Microsoft 365 Business Premium"/>
                            <ComboBoxItem Content="Office 365 E3"/>
                            <ComboBoxItem Content="Office 365 E5"/>
                        </ComboBox>
                    </StackPanel>
                </CheckBox>
                
                <CheckBox x:Name="chkDryRun" Style="{StaticResource ModernCheckBox}" 
                         Content="Dry Run (preview only - don't create users)" IsChecked="True" Foreground="#FFFFA500"/>
            </StackPanel>
        </Border>
        
        <!-- Results -->
        <Border Grid.Row="4" Style="{StaticResource CardBorder}" MaxHeight="200">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="Results" FontSize="14" FontWeight="SemiBold" Foreground="White"/>
                    <Button x:Name="btnSaveLog" Content="Save Log" Margin="10,0,0,0" Padding="10,3"/>
                    <Button x:Name="btnExportResults" Content="Export Results" Margin="10,0,0,0" Padding="10,3"/>
                    <Button x:Name="btnClear" Content="Clear" Margin="10,0,0,0" Padding="10,3"/>
                </StackPanel>
                
                <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                    <TextBox x:Name="txtResults" Background="Transparent" Foreground="White" 
                            BorderThickness="0" IsReadOnly="True" FontFamily="Consolas" 
                            TextWrapping="Wrap" AcceptsReturn="True"/>
                </ScrollViewer>
            </Grid>
        </Border>
        
        <!-- Progress -->
        <Grid Grid.Row="5" Margin="0,10,0,0">
            <ProgressBar x:Name="progressBar" Height="20" Minimum="0" Maximum="100"/>
            <TextBlock x:Name="txtProgress" HorizontalAlignment="Center" VerticalAlignment="Center" Foreground="White"/>
        </Grid>
        
        <!-- Action Buttons -->
        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,40,0,0">
            <Button x:Name="btnValidate" Style="{StaticResource ModernButton}" Content="Validate Data" Width="120" Margin="0,0,10,0"/>
            
            <Button x:Name="btnPreview" Style="{StaticResource ModernButton}" Content="Preview Data" Width="120" Margin="0,0,10,0"/>
            
            <Button x:Name="btnImport" Style="{StaticResource ModernButton}" Content="Import Users" Width="120" Background="#FF107C10"/>
        </StackPanel>
    </Grid>
</Window>
"@
#endregion

#region Functions

function Connect-M365 {
    param($Ui)
    
    if ($script:Connected) {
        try {
            Disconnect-MgGraph | Out-Null
            $script:Connected = $false
            Update-Status -Ui $Ui -Connected $false
            $Ui.TxtResults.Text = "Disconnected from Microsoft 365"
        }
        catch {
            $Ui.TxtResults.Text = "Error disconnecting: $_"
        }
        return
    }
    
    try {
        Connect-MgGraph -Scopes @(
            "User.ReadWrite.All"
            "Directory.ReadWrite.All"
            "Organization.Read.All"
            "Group.Read.All"
            "GroupMember.ReadWrite.All"
        ) -ErrorAction Stop
        
        $context = Get-MgContext
        $script:Connected = $true
        Update-Status -Ui $Ui -Connected $true -Account $context.Account
        $Ui.TxtResults.Text = "Connected to Microsoft 365 as $($context.Account)"
    }
    catch {
        $Ui.TxtResults.Text = "Failed to connect: $_"
        [System.Windows.MessageBox]::Show("Failed to connect: $_", "Error", "OK", "Error")
    }
}

function Update-Status {
    param($Ui, [bool]$Connected, [string]$Account = "")
    
    if ($Connected) {
        $Ui.StatusIndicator.Fill = "#FF107C10"
        $Ui.StatusText.Text = "Connected"
        $Ui.StatusAccount.Text = $Account
        $Ui.BtnConnect.Content = "Disconnect"
    }
    else {
        $Ui.StatusIndicator.Fill = "#FFE81123"
        $Ui.StatusText.Text = "Not Connected"
        $Ui.StatusAccount.Text = ""
        $Ui.BtnConnect.Content = "Connect to M365"
    }
}

function Test-EmailFormat {
    param([string]$Email)
    return $Email -match '^[\w\.-]+@[\w\.-]+\.\w+$'
}

function Test-PasswordStrength {
    param([string]$Password)
    
    $score = 0
    if ($Password.Length -ge 8) { $score++ }
    if ($Password -match '[A-Z]') { $score++ }
    if ($Password -match '[a-z]') { $score++ }
    if ($Password -match '\d') { $score++ }
    if ($Password -match '[!@#$%^&*(),.?":{}|<>]') { $score++ }
    
    return $score -ge 3
}

function Validate-UserData {
    param($Users, $Ui, [bool]$CheckExisting)
    
    $results = @()
    $validCount = 0
    $errorCount = 0
    
    $Ui.TxtResults.Text = "Validating data..."
    $Ui.ProgressBar.Maximum = $Users.Count
    $Ui.ProgressBar.Value = 0
    
    # Get existing users if needed
    $existingUsers = @()
    if ($CheckExisting -and $script:Connected) {
        $Ui.TxtResults.Text += "`nLoading existing users from Azure AD..."
        $existingUsers = Get-MgUser -All | Select-Object -ExpandProperty UserPrincipalName
    }
    
    for ($i = 0; $i -lt $Users.Count; $i++) {
        $user = $Users[$i]
        $lineNum = $i + 2  # Excel row number (header + 1-based)
        $errors = @()
        $warnings = @()
        
        # Required fields
        if ([string]::IsNullOrWhiteSpace($user.DisplayName)) {
            $errors += "Missing DisplayName"
        }
        if ([string]::IsNullOrWhiteSpace($user.UserPrincipalName)) {
            $errors += "Missing UserPrincipalName"
        }
        if ([string]::IsNullOrWhiteSpace($user.MailNickname)) {
            $errors += "Missing MailNickname"
        }
        if ([string]::IsNullOrWhiteSpace($user.Password)) {
            $errors += "Missing Password"
        }
        
        # Email format validation
        if ($user.UserPrincipalName -and -not (Test-EmailFormat $user.UserPrincipalName)) {
            $errors += "Invalid email format: $($user.UserPrincipalName)"
        }
        
        # Password strength
        if ($user.Password -and -not (Test-PasswordStrength $user.Password)) {
            $warnings += "Weak password (use 8+ chars with mixed case, numbers, symbols)"
        }
        
        # Check for duplicates in the Excel file
        $duplicateUPN = $Users | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }
        if ($duplicateUPN.Count -gt 1) {
            $errors += "Duplicate UserPrincipalName in Excel file"
        }
        
        # Check if user already exists in Azure AD
        if ($CheckExisting -and $existingUsers -contains $user.UserPrincipalName) {
            $errors += "User already exists in Azure AD"
        }
        
        # Build result
        if ($errors.Count -gt 0) {
            $results += "Row $lineNum : $($user.DisplayName) [$($user.UserPrincipalName)]"
            $results += "   ERRORS: $($errors -join ', ')"
            if ($warnings.Count -gt 0) {
                $results += "   WARNINGS: $($warnings -join ', ')"
            }
            $results += ""
            $errorCount++
        }
        elseif ($warnings.Count -gt 0) {
            $results += "Row $lineNum : $($user.DisplayName) [$($user.UserPrincipalName)]"
            $results += "   WARNINGS: $($warnings -join ', ')"
            $results += ""
            $validCount++
        }
        else {
            $validCount++
        }
        
        $Ui.ProgressBar.Value = $i + 1
        $Ui.TxtProgress.Text = "Validating: $($i + 1) of $($Users.Count)"
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    $Ui.TxtProgress.Text = "Validation Complete"
    
    $summary = @"
================================
VALIDATION SUMMARY
================================
Total: $($Users.Count)
Valid: $validCount
Errors: $errorCount
================================

"@
    
    if ($results.Count -gt 0) {
        return $summary + ($results -join "`n")
    }
    else {
        return $summary + "All users passed validation!"
    }
}

function Import-UsersFromExcel {
    param($Ui, [bool]$DryRun, [bool]$ValidateFirst)
    
    if (-not $script:Connected) {
        [System.Windows.MessageBox]::Show("Please connect to Microsoft 365 first.", "Not Connected", "OK", "Warning")
        return
    }
    
    if (-not $script:ImportData) {
        [System.Windows.MessageBox]::Show("Please select and validate an Excel file first.", "No Data", "OK", "Warning")
        return
    }
    
    # Run validation first if requested
    if ($ValidateFirst) {
        $validationResult = Validate-UserData -Users $script:ImportData -Ui $Ui -CheckExisting $Ui.ChkCheckExisting.IsChecked
        $Ui.TxtResults.Text = $validationResult
        
        if ($validationResult -match "ERRORS") {
            $continue = [System.Windows.MessageBox]::Show("Validation found errors. Continue anyway?", "Validation Errors", "YesNo", "Warning")
            if ($continue -ne "Yes") { return }
        }
    }
    
    $forcePasswordChange = $Ui.ChkForcePasswordChange.IsChecked
    $assignLicense = $Ui.ChkAssignLicense.IsChecked
    $selectedLicense = $Ui.CmbLicense.SelectedItem
    
    $mode = if ($DryRun) { "DRY RUN" } else { "LIVE IMPORT" }
    $icon = if ($DryRun) { "Information" } else { "Warning" }
    
    $confirm = [System.Windows.MessageBox]::Show(
        "About to import $($script:ImportData.Count) user(s).`n`nMode: $mode`n`nContinue?",
        "Confirm Import", "YesNo", $icon)
    
    if ($confirm -ne "Yes") { return }
    
    $successCount = 0
    $failCount = 0
    $results = @()
    $script:ImportLog = @()
    
    $Ui.ProgressBar.Maximum = $script:ImportData.Count
    $Ui.ProgressBar.Value = 0
    
    foreach ($user in $script:ImportData) {
        $logEntry = @{
            Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            DisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            Status = ""
            Message = ""
        }
        
        try {
            $Ui.TxtProgress.Text = "Processing $($successCount + $failCount + 1) of $($script:ImportData.Count)"
            
            if ($DryRun) {
                $msg = "[DRY RUN] Would create: $($user.DisplayName) ($($user.UserPrincipalName))"
                $results += $msg
                $logEntry.Status = "DRY_RUN"
                $logEntry.Message = $msg
                $successCount++
            }
            else {
                $passwordProfile = @{
                    Password = $user.Password
                    ForceChangePasswordNextSignIn = $forcePasswordChange
                }
                
                $userParams = @{
                    DisplayName = $user.DisplayName
                    UserPrincipalName = $user.UserPrincipalName
                    MailNickname = $user.MailNickname
                    AccountEnabled = $true
                    PasswordProfile = $passwordProfile
                }
                
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
                
                $newUser = New-MgUser @userParams
                
                if ($assignLicense -and $selectedLicense) {
                    $skuPartNumber = switch ($selectedLicense.Content) {
                        "Microsoft 365 Business Basic" { "O365_BUSINESS_ESSENTIALS" }
                        "Microsoft 365 Business Standard" { "O365_BUSINESS_PREMIUM" }
                        "Microsoft 365 Business Premium" { "SPB" }
                        "Office 365 E3" { "ENTERPRISEPACK" }
                        "Office 365 E5" { "ENTERPRISEPREMIUM" }
                    }
                    
                    if ($skuPartNumber) {
                        $sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $skuPartNumber }
                        if ($sku) {
                            $license = @{ AddLicenses = @(@{ SkuId = $sku.SkuId }); RemoveLicenses = @() }
                            Set-MgUserLicense -UserId $newUser.Id -BodyParameter $license | Out-Null
                            $results += "  [LICENSE] Assigned: $($selectedLicense.Content)"
                        }
                    }
                }
                
                if ($user.Groups) {
                    $groupList = $user.Groups -split ";"
                    foreach ($groupName in $groupList) {
                        $group = Get-MgGroup -Filter "displayName eq '$($groupName.Trim())'" -ErrorAction SilentlyContinue
                        if ($group) {
                            New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $newUser.Id -ErrorAction SilentlyContinue | Out-Null
                            $results += "  [GROUP] Added to: $($groupName.Trim())"
                        }
                    }
                }
                
                $msg = "Created: $($user.DisplayName) ($($user.UserPrincipalName))"
                $results += $msg
                $logEntry.Status = "SUCCESS"
                $logEntry.Message = $msg
                $successCount++
            }
        }
        catch {
            $msg = "FAILED: $($user.DisplayName) - $($_.Exception.Message)"
            $results += $msg
            $logEntry.Status = "FAILED"
            $logEntry.Message = $_.Exception.Message
            $failCount++
        }
        
        $script:ImportLog += $logEntry
        $Ui.ProgressBar.Value++
        $Ui.TxtResults.Text = $results -join "`n"
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    $Ui.TxtProgress.Text = "Complete! Success: $successCount | Failed: $failCount"
    $icon2 = if ($failCount -eq 0) { "Information" } else { "Warning" }
    [System.Windows.MessageBox]::Show("Import complete!`n`nSuccessful: $successCount`nFailed: $failCount", "Done", "OK", $icon2)
}

function Export-Template {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
    $saveDialog.FileName = "M365-Import-Template.xlsx"
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $template = @(
            [PSCustomObject]@{
                DisplayName = "John Doe"
                UserPrincipalName = "john.doe@company.com"
                MailNickname = "johndoe"
                GivenName = "John"
                Surname = "Doe"
                JobTitle = "IT Administrator"
                Department = "IT"
                OfficeLocation = "Amsterdam"
                MobilePhone = "+31-6-12345678"
                BusinessPhones = "+31-20-1234567"
                StreetAddress = "Hoofdstraat 123"
                City = "Amsterdam"
                State = "Noord-Holland"
                PostalCode = "1000 AA"
                Country = "Nederland"
                UsageLocation = "NL"
                Password = "Welcome2024!"
                Groups = "IT Team;All Employees"
            },
            [PSCustomObject]@{
                DisplayName = "Jane Smith"
                UserPrincipalName = "jane.smith@company.com"
                MailNickname = "janesmith"
                GivenName = "Jane"
                Surname = "Smith"
                JobTitle = "HR Manager"
                Department = "HR"
                OfficeLocation = "Rotterdam"
                MobilePhone = "+31-6-87654321"
                BusinessPhones = "+31-10-7654321"
                StreetAddress = "Stationplein 45"
                City = "Rotterdam"
                State = "Zuid-Holland"
                PostalCode = "3013 AK"
                Country = "Nederland"
                UsageLocation = "NL"
                Password = "Welcome2024!"
                Groups = "HR Team;All Employees"
            }
        )
        
        $template | Export-Excel -Path $saveDialog.FileName -AutoSize -TableName "Users"
        [System.Windows.MessageBox]::Show("Template saved to:`n$($saveDialog.FileName)", "Template Created", "OK", "Information")
    }
}

#endregion

#region Main

$script:Connected = $false
$script:ImportData = $null
$script:ImportLog = @()
$script:FilePath = $null

$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# UI Elements
$Ui = @{
    Window = $Window
    StatusIndicator = $Window.FindName("statusIndicator")
    StatusText = $Window.FindName("statusText")
    StatusAccount = $Window.FindName("statusAccount")
    BtnConnect = $Window.FindName("btnConnect")
    TxtFilePath = $Window.FindName("txtFilePath")
    BtnBrowse = $Window.FindName("btnBrowse")
    BtnDownloadTemplate = $Window.FindName("btnDownloadTemplate")
    ChkValidateData = $Window.FindName("chkValidateData")
    ChkCheckExisting = $Window.FindName("chkCheckExisting")
    ChkForcePasswordChange = $Window.FindName("chkForcePasswordChange")
    ChkAssignLicense = $Window.FindName("chkAssignLicense")
    CmbLicense = $Window.FindName("cmbLicense")
    ChkDryRun = $Window.FindName("chkDryRun")
    TxtResults = $Window.FindName("txtResults")
    BtnSaveLog = $Window.FindName("btnSaveLog")
    BtnExportResults = $Window.FindName("btnExportResults")
    BtnClear = $Window.FindName("btnClear")
    ProgressBar = $Window.FindName("progressBar")
    TxtProgress = $Window.FindName("txtProgress")
    BtnValidate = $Window.FindName("btnValidate")
    BtnPreview = $Window.FindName("btnPreview")
    BtnImport = $Window.FindName("btnImport")
}

# Events
$Ui.BtnConnect.Add_Click({ Connect-M365 -Ui $Ui })

$Ui.BtnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($dialog.ShowDialog() -eq "OK") {
        $Ui.TxtFilePath.Text = $dialog.FileName
        $script:FilePath = $dialog.FileName
        
        # Auto load data
        try {
            $script:ImportData = Import-Excel -Path $script:FilePath
            $Ui.TxtResults.Text = "Loaded $($script:ImportData.Count) users from file.`nClick 'Validate Data' to check for errors."
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to load file: $_", "Error", "OK", "Error")
        }
    }
})

$Ui.BtnDownloadTemplate.Add_Click({ Export-Template })

$Ui.BtnValidate.Add_Click({
    if (-not $script:ImportData) {
        [System.Windows.MessageBox]::Show("Please select an Excel file first.", "No Data", "OK", "Warning")
        return
    }
    
    $validationResult = Validate-UserData -Users $script:ImportData -Ui $Ui -CheckExisting $Ui.ChkCheckExisting.IsChecked
    $Ui.TxtResults.Text = $validationResult
})

$Ui.BtnPreview.Add_Click({
    if (-not $script:ImportData) {
        [System.Windows.MessageBox]::Show("Please select an Excel file first.", "No Data", "OK", "Warning")
        return
    }
    
    $Ui.TxtResults.Text = "Preview of users to import:`n`n"
    foreach ($user in $script:ImportData | Select-Object -First 10) {
        $Ui.TxtResults.Text += "• $($user.DisplayName) [$($user.UserPrincipalName)] - $($user.Department)`n"
    }
    if ($script:ImportData.Count -gt 10) {
        $Ui.TxtResults.Text += "`n... and $($script:ImportData.Count - 10) more users"
    }
})

$Ui.BtnImport.Add_Click({
    Import-UsersFromExcel -Ui $Ui -DryRun $Ui.ChkDryRun.IsChecked -ValidateFirst $Ui.ChkValidateData.IsChecked
})

$Ui.BtnSaveLog.Add_Click({
    if ($script:ImportLog.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No import log to save.", "No Data", "OK", "Warning")
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "Import-Log-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $script:ImportLog | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
        [System.Windows.MessageBox]::Show("Log saved to:`n$($saveDialog.FileName)", "Log Saved", "OK", "Information")
    }
})

$Ui.BtnExportResults.Add_Click({
    $text = $Ui.TxtResults.Text
    if ([string]::IsNullOrWhiteSpace($text)) {
        [System.Windows.MessageBox]::Show("No results to export.", "No Data", "OK", "Warning")
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text files (*.txt)|*.txt"
    $saveDialog.FileName = "Import-Results-$(Get-Date -Format 'yyyyMMdd-HHmmss').txt"
    $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $text | Out-File -FilePath $saveDialog.FileName
        [System.Windows.MessageBox]::Show("Results saved to:`n$($saveDialog.FileName)", "Results Saved", "OK", "Information")
    }
})

$Ui.BtnClear.Add_Click({
    $Ui.TxtResults.Text = ""
    $Ui.ProgressBar.Value = 0
    $Ui.TxtProgress.Text = ""
})

# Show Window
$Window.ShowDialog() | Out-Null

#endregion
