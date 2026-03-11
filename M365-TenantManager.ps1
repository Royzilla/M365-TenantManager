<#
.SYNOPSIS
    M365 Tenant Manager - Modern GUI for Microsoft 365 administration

.DESCRIPTION
    Professional PowerShell WPF application for managing Microsoft 365 tenants
    via Microsoft Graph API. Features bulk user creation, license management,
    and group administration.

.AUTHOR
    M365 Tenant Manager

.VERSION
    1.0.0
#>

#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.DirectoryManagement, ImportExcel

#region Initialization
param(
    [switch]$DebugMode
)

# Error handling
$ErrorActionPreference = "Stop"
$ProgressPreference = "Continue"

# Add required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Import modules
$script:ModuleRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$script:ModuleRoot\Modules\Graph-Auth.ps1"
. "$script:ModuleRoot\Modules\User-Management.ps1"
. "$script:ModuleRoot\Modules\License-Management.ps1"
. "$script:ModuleRoot\Modules\Group-Management.ps1"
. "$script:ModuleRoot\Modules\UI-Helpers.ps1"

# Global state
$script:GraphConnection = @{
    Connected = $false
    Account = $null
    Tenant = $null
    TokenExpiry = $null
}

$script:CurrentTab = "Dashboard"
#endregion

#region XAML UI Definition
[xml]$XAML = @"
&lt;Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Tenant Manager" Height="800" Width="1200"
        WindowStartupLocation="CenterScreen"
        Background="&#35;FF1E1E1E"
        Foreground="White"
        FontFamily="Segoe UI"
        MinWidth="1000" MinHeight="700"&gt;
    
    &lt;Window.Resources&gt;
        &lt;Style x:Key="ModernButton" TargetType="Button"&gt;
            &lt;Setter Property="Background" Value="&#35;FF0078D4"/&gt;
            &lt;Setter Property="Foreground" Value="White"/&gt;
            &lt;Setter Property="BorderThickness" Value="0"/&gt;
            &lt;Setter Property="Padding" Value="15,8"/&gt;
            &lt;Setter Property="FontSize" Value="13"/&gt;
            &lt;Setter Property="Cursor" Value="Hand"/&gt;
            &lt;Setter Property="Template"&gt;
                &lt;Setter.Value&gt;
                    &lt;ControlTemplate TargetType="Button"&gt;
                        &lt;Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                BorderThickness="0"&gt;
                            &lt;ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/&gt;
                        &lt;/Border&gt;
                    &lt;/ControlTemplate&gt;
                &lt;/Setter.Value&gt;
            &lt;/Setter&gt;
            &lt;Style.Triggers&gt;
                &lt;Trigger Property="IsMouseOver" Value="True"&gt;
                    &lt;Setter Property="Background" Value="&#35;FF106EBE"/&gt;
                &lt;/Trigger&gt;
                &lt;Trigger Property="IsEnabled" Value="False"&gt;
                    &lt;Setter Property="Background" Value="&#35;FF6C6C6C"/&gt;
                &lt;/Trigger&gt;
            &lt;/Style.Triggers&gt;
        &lt;/Style&gt;
        
        &lt;Style x:Key="NavButton" TargetType="Button"&gt;
            &lt;Setter Property="Background" Value="Transparent"/&gt;
            &lt;Setter Property="Foreground" Value="&#35;FFCCCCCC"/&gt;
            &lt;Setter Property="BorderThickness" Value="0"/&gt;
            &lt;Setter Property="Padding" Value="15,12"/&gt;
            &lt;Setter Property="FontSize" Value="14"/&gt;
            &lt;Setter Property="HorizontalContentAlignment" Value="Left"/&gt;
            &lt;Setter Property="Template"&gt;
                &lt;Setter.Value&gt;
                    &lt;ControlTemplate TargetType="Button"&gt;
                        &lt;Border Background="{TemplateBinding Background}" 
                                BorderThickness="4,0,0,0"
                                BorderBrush="Transparent"&gt;
                            &lt;ContentPresenter Margin="10,0,0,0" 
                                            VerticalAlignment="Center"/&gt;
                        &lt;/Border&gt;
                    &lt;/ControlTemplate&gt;
                &lt;/Setter.Value&gt;
            &lt;/Setter&gt;
            &lt;Style.Triggers&gt;
                &lt;Trigger Property="IsMouseOver" Value="True"&gt;
                    &lt;Setter Property="Background" Value="&#35;FF2D2D2D"/&gt;
                    &lt;Setter Property="Foreground" Value="White"/&gt;
                &lt;/Trigger&gt;
            &lt;/Style.Triggers&gt;
        &lt;/Style&gt;
        
        &lt;Style x:Key="CardBorder" TargetType="Border"&gt;
            &lt;Setter Property="Background" Value="&#35;FF252526"/&gt;
            &lt;Setter Property="BorderBrush" Value="&#35;FF3E3E42"/&gt;
            &lt;Setter Property="BorderThickness" Value="1"/&gt;
            &lt;Setter Property="CornerRadius" Value="8"/&gt;
            &lt;Setter Property="Padding" Value="20"/&gt;
            &lt;Setter Property="Margin" Value="0,0,0,15"/&gt;
        &lt;/Style&gt;
        
        &lt;Style x:Key="ModernTextBox" TargetType="TextBox"&gt;
            &lt;Setter Property="Background" Value="&#35;FF3C3C3C"/&gt;
            &lt;Setter Property="Foreground" Value="White"/&gt;
            &lt;Setter Property="BorderBrush" Value="&#35;FF555555"/&gt;
            &lt;Setter Property="BorderThickness" Value="1"/&gt;
            &lt;Setter Property="Padding" Value="10,8"/&gt;
            &lt;Setter Property="FontSize" Value="13"/&gt;
        &lt;/Style&gt;
        
        &lt;Style x:Key="ModernComboBox" TargetType="ComboBox"&gt;
            &lt;Setter Property="Background" Value="&#35;FF3C3C3C"/&gt;
            &lt;Setter Property="Foreground" Value="Black"/&gt;
            &lt;Setter Property="BorderBrush" Value="&#35;FF555555"/&gt;
            &lt;Setter Property="Padding" Value="10,8"/&gt;
            &lt;Setter Property="FontSize" Value="13"/&gt;
        &lt;/Style&gt;
        
        &lt;Style x:Key="ModernCheckBox" TargetType="CheckBox"&gt;
            &lt;Setter Property="Foreground" Value="White"/&gt;
            &lt;Setter Property="FontSize" Value="13"/&gt;
            &lt;Setter Property="Margin" Value="0,5"/&gt;
        &lt;/Style&gt;
    &lt;/Window.Resources&gt;
    
    &lt;Grid&gt;
        &lt;Grid.ColumnDefinitions&gt;
            &lt;ColumnDefinition Width="250"/&gt;
            &lt;ColumnDefinition Width="*"/&gt;
        &lt;/Grid.ColumnDefinitions&gt;
        
        &lt;!-- Sidebar --&gt;
        &lt;Border Grid.Column="0" Background="&#35;FF252526" BorderBrush="&#35;FF3E3E42" BorderThickness="0,0,1,0"&gt;
            &lt;Grid&gt;
                &lt;Grid.RowDefinitions&gt;
                    &lt;RowDefinition Height="Auto"/&gt;
                    &lt;RowDefinition Height="Auto"/&gt;
                    &lt;RowDefinition Height="*"/&gt;
                    &lt;RowDefinition Height="Auto"/&gt;
                &lt;/Grid.RowDefinitions&gt;
                
                &lt;!-- Logo/Title --&gt;
                &lt;Border Grid.Row="0" Padding="20,25"&gt;
                    &lt;StackPanel&gt;
                        &lt;TextBlock Text="M365" FontSize="28" FontWeight="Bold" Foreground="&#35;FF0078D4"/&gt;
                        &lt;TextBlock Text="Tenant Manager" FontSize="14" Foreground="&#35;FFCCCCCC" Margin="0,2,0,0"/&gt;
                    &lt;/StackPanel&gt;
                &lt;/Border&gt;
                
                &lt;!-- Connection Status --&gt;
                &lt;Border Grid.Row="1" Background="&#35;FF1E1E1E" Padding="15" Margin="10,0" CornerRadius="4"&gt;
                    &lt;StackPanel&gt;
                        &lt;TextBlock Text="Connection Status" FontSize="11" Foreground="&#35;FF858585" Margin="0,0,0,5"/&gt;
                        &lt;StackPanel Orientation="Horizontal"&gt;
                            &lt;Ellipse x:Name="statusIndicator" Width="10" Height="10" Fill="&#35;FFE81123" Margin="0,0,8,0"/&gt;
                            &lt;TextBlock x:Name="statusText" Text="Disconnected" FontSize="12" Foreground="White"/&gt;
                        &lt;/StackPanel&gt;
                        &lt;TextBlock x:Name="statusAccount" Text="" FontSize="10" Foreground="&#35;FF858585" Margin="18,3,0,0" TextTrimming="CharacterEllipsis"/&gt;
                    &lt;/StackPanel&gt;
                &lt;/Border&gt;
                
                &lt;!-- Navigation --&gt;
                &lt;StackPanel Grid.Row="2" Margin="0,20,0,0"&gt;
                    &lt;Button x:Name="navDashboard" Style="{StaticResource NavButton}" Content="📊 Dashboard"/&gt;
                    &lt;Button x:Name="navUsers" Style="{StaticResource NavButton}" Content="👤 User Management"/&gt;
                    &lt;Button x:Name="navLicenses" Style="{StaticResource NavButton}" Content="🔑 Licenses"/&gt;
                    &lt;Button x:Name="navGroups" Style="{StaticResource NavButton}" Content="👥 Groups"/&gt;
                    &lt;Button x:Name="navReports" Style="{StaticResource NavButton}" Content="📈 Reports"/&gt;
                    &lt;Button x:Name="navSettings" Style="{StaticResource NavButton}" Content="⚙️ Settings"/&gt;
                &lt;/StackPanel&gt;
                
                &lt;!-- Version --&gt;
                &lt;TextBlock Grid.Row="3" Text="v1.0.0" FontSize="11" Foreground="&#35;FF666666" 
                          HorizontalAlignment="Center" Margin="0,15"/&gt;
            &lt;/Grid&gt;
        &lt;/Border&gt;
        
        &lt;!-- Main Content --&gt;
        &lt;Grid Grid.Column="1" Background="&#35;FF1E1E1E"&gt;
            &lt;Grid.RowDefinitions&gt;
                &lt;RowDefinition Height="Auto"/&gt;
                &lt;RowDefinition Height="*"/&gt;
            &lt;/Grid.RowDefinitions&gt;
            
            &lt;!-- Header --&gt;
            &lt;Border Grid.Row="0" Background="&#35;FF252526" BorderBrush="&#35;FF3E3E42" BorderThickness="0,0,0,1" Padding="25,20"&gt;
                &lt;Grid&gt;
                    &lt;Grid.ColumnDefinitions&gt;
                        &lt;ColumnDefinition Width="*"/&gt;
                        &lt;ColumnDefinition Width="Auto"/&gt;
                    &lt;/Grid.ColumnDefinitions&gt;
                    &lt;TextBlock x:Name="pageTitle" Grid.Column="0" Text="Dashboard" FontSize="24" FontWeight="SemiBold" Foreground="White"/&gt;
                    &lt;Button x:Name="btnConnect" Grid.Column="1" Style="{StaticResource ModernButton}" 
                           Content="Connect to M365" Padding="20,10"/&gt;
                &lt;/Grid&gt;
            &lt;/Border&gt;
            
            &lt;!-- Content Area --&gt;
            &lt;ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"&gt;
                &lt;ContentControl x:Name="contentArea" Margin="25"&gt;
                    
                    &lt;!-- Dashboard View (Default) --&gt;
                    &lt;Grid x:Name="viewDashboard"&gt;
                        &lt;Grid.RowDefinitions&gt;
                            &lt;RowDefinition Height="Auto"/&gt;
                            &lt;RowDefinition Height="Auto"/&gt;
                        &lt;/Grid.RowDefinitions&gt;
                        
                        &lt;!-- Stats Cards --&gt;
                        &lt;UniformGrid Grid.Row="0" Columns="4" Margin="0,0,0,20"&gt;
                            &lt;Border Style="{StaticResource CardBorder}"&gt;
                                &lt;StackPanel&gt;
                                    &lt;TextBlock Text="👤" FontSize="24" Margin="0,0,0,10"/&gt;
                                    &lt;TextBlock x:Name="statTotalUsers" Text="-" FontSize="32" FontWeight="Bold" Foreground="White"/&gt;
                                    &lt;TextBlock Text="Total Users" FontSize="12" Foreground="&#35;FF858585" Margin="0,5,0,0"/&gt;
                                &lt;/StackPanel&gt;
                            &lt;/Border&gt;
                            
                            &lt;Border Style="{StaticResource CardBorder}"&gt;
                                &lt;StackPanel&gt;
                                    &lt;TextBlock Text="✅" FontSize="24" Margin="0,0,0,10"/&gt;
                                    &lt;TextBlock x:Name="statActiveUsers" Text="-" FontSize="32" FontWeight="Bold" Foreground="&#35;FF107C10"/&gt;
                                    &lt;TextBlock Text="Active Users" FontSize="12" Foreground="&#35;FF858585" Margin="0,5,0,0"/&gt;
                                &lt;/StackPanel&gt;
                            &lt;/Border&gt;
                            
                            &lt;Border Style="{StaticResource CardBorder}"&gt;
                                &lt;StackPanel&gt;
                                    &lt;TextBlock Text="🔑" FontSize="24" Margin="0,0,0,10"/&gt;
                                    &lt;TextBlock x:Name="statLicenses" Text="-" FontSize="32" FontWeight="Bold" Foreground="&#35;FFFFA500"/&gt;
                                    &lt;TextBlock Text="Licenses Used" FontSize="12" Foreground="&#35;FF858585" Margin="0,5,0,0"/&gt;
                                &lt;/StackPanel&gt;
                            &lt;/Border&gt;
                            
                            &lt;Border Style="{StaticResource CardBorder}"&gt;
                                &lt;StackPanel&gt;
                                    &lt;TextBlock Text="👥" FontSize="24" Margin="0,0,0,10"/&gt;
                                    &lt;TextBlock x:Name="statGroups" Text="-" FontSize="32" FontWeight="Bold" Foreground="&#35;FF0078D4"/&gt;
                                    &lt;TextBlock Text="Groups" FontSize="12" Foreground="&#35;FF858585" Margin="0,5,0,0"/&gt;
                                &lt;/StackPanel&gt;
                            &lt;/Border&gt;
                        &lt;/UniformGrid&gt;
                        
                        &lt;!-- Quick Actions --&gt;
                        &lt;Border Grid.Row="1" Style="{StaticResource CardBorder}"&gt;
                            &lt;StackPanel&gt;
                                &lt;TextBlock Text="Quick Actions" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,20"/&gt;
                                &lt;WrapPanel&gt;
                                    &lt;Button x:Name="btnQuickBulkImport" Style="{StaticResource ModernButton}" 
                                           Content="📊 Bulk Import Users" Margin="0,0,10,10" Width="180"/&gt;
                                    &lt;Button x:Name="btnQuickCreateUser" Style="{StaticResource ModernButton}" 
                                           Content="➕ Create Single User" Margin="0,0,10,10" Width="180"/&gt;
                                    &lt;Button x:Name="btnQuickAssignLicense" Style="{StaticResource ModernButton}" 
                                           Content="🔑 Assign Licenses" Margin="0,0,10,10" Width="180"/&gt;
                                    &lt;Button x:Name="btnQuickExport" Style="{StaticResource ModernButton}" 
                                           Content="📤 Export Users" Margin="0,0,10,10" Width="180"/&gt;
                                &lt;/WrapPanel&gt;
                            &lt;/StackPanel&gt;
                        &lt;/Border&gt;
                    &lt;/Grid&gt;
                    
                &lt;/ContentControl&gt;
            &lt;/ScrollViewer&gt;
        &lt;/Grid&gt;
    &lt;/Grid&gt;
&lt;/Window&gt;
"@
#endregion

#region Create and Show Window
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Get UI elements
$script:ui = @{
    Window = $Window
    StatusIndicator = $Window.FindName("statusIndicator")
    StatusText = $Window.FindName("statusText")
    StatusAccount = $Window.FindName("statusAccount")
    PageTitle = $Window.FindName("pageTitle")
    ContentArea = $Window.FindName("contentArea")
    BtnConnect = $Window.FindName("btnConnect")
    
    # Stats
    StatTotalUsers = $Window.FindName("statTotalUsers")
    StatActiveUsers = $Window.FindName("statActiveUsers")
    StatLicenses = $Window.FindName("statLicenses")
    StatGroups = $Window.FindName("statGroups")
    
    # Quick Actions
    BtnQuickBulkImport = $Window.FindName("btnQuickBulkImport")
    BtnQuickCreateUser = $Window.FindName("btnQuickCreateUser")
    BtnQuickAssignLicense = $Window.FindName("btnQuickAssignLicense")
    BtnQuickExport = $Window.FindName("btnQuickExport")
    
    # Navigation
    NavDashboard = $Window.FindName("navDashboard")
    NavUsers = $Window.FindName("navUsers")
    NavLicenses = $Window.FindName("navLicenses")
    NavGroups = $Window.FindName("navGroups")
    NavReports = $Window.FindName("navReports")
    NavSettings = $Window.FindName("navSettings")
}

# Initialize UI
Update-ConnectionStatus -Ui $script:ui -Status "Disconnected"

# Wire up events
$script:ui.BtnConnect.Add_Click({ Connect-M365 -Ui $script:ui })
$script:ui.BtnQuickBulkImport.Add_Click({ Show-BulkImportView -Ui $script:ui })
$script:ui.BtnQuickCreateUser.Add_Click({ Show-CreateUserView -Ui $script:ui })
$script:ui.NavDashboard.Add_Click({ Show-DashboardView -Ui $script:ui })
$script:ui.NavUsers.Add_Click({ Show-UserManagementView -Ui $script:ui })
$script:ui.NavLicenses.Add_Click({ Show-LicenseView -Ui $script:ui })
$script:ui.NavGroups.Add_Click({ Show-GroupView -Ui $script:ui })

# Show window
$Window.ShowDialog() | Out-Null
#endregion
