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
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="M365 Tenant Manager" Height="800" Width="1200"
        WindowStartupLocation="CenterScreen"
        Background="#FF1E1E1E"
        Foreground="White"
        FontFamily="Segoe UI"
        MinWidth="1000" MinHeight="700">
    
    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#FF0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="4" 
                                BorderThickness="0">
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
        
        <Style x:Key="NavButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#FFCCCCCC"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,12"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderThickness="4,0,0,0"
                                BorderBrush="Transparent">
                            <ContentPresenter Margin="10,0,0,0" 
                                            VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF2D2D2D"/>
                    <Setter Property="Foreground" Value="White"/>
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
            <Setter Property="Background" Value="#FF3C3C3C"/>
            <Setter Property="Foreground" Value="Black"/>
            <Setter Property="BorderBrush" Value="#FF555555"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        
        <Style x:Key="ModernCheckBox" TargetType="CheckBox">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,5"/>
        </Style>
    </Window.Resources>
    
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="250"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        
        <!-- Sidebar -->
        <Border Grid.Column="0" Background="#FF252526" BorderBrush="#FF3E3E42" BorderThickness="0,0,1,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <!-- Logo/Title -->
                <Border Grid.Row="0" Padding="20,25">
                    <StackPanel>
                        <TextBlock Text="M365" FontSize="28" FontWeight="Bold" Foreground="#FF0078D4"/>
                        <TextBlock Text="Tenant Manager" FontSize="14" Foreground="#FFCCCCCC" Margin="0,2,0,0"/>
                    </StackPanel>
                </Border>
                
                <!-- Connection Status -->
                <Border Grid.Row="1" Background="#FF1E1E1E" Padding="15" Margin="10,0" CornerRadius="4">
                    <StackPanel>
                        <TextBlock Text="Connection Status" FontSize="11" Foreground="#FF858585" Margin="0,0,0,5"/>
                        <StackPanel Orientation="Horizontal">
                            <Ellipse x:Name="statusIndicator" Width="10" Height="10" Fill="#FFE81123" Margin="0,0,8,0"/>
                            <TextBlock x:Name="statusText" Text="Disconnected" FontSize="12" Foreground="White"/>
                        </StackPanel>
                        <TextBlock x:Name="statusAccount" Text="" FontSize="10" Foreground="#FF858585" Margin="18,3,0,0" TextTrimming="CharacterEllipsis"/>
                    </StackPanel>
                </Border>
                
                <!-- Navigation -->
                <StackPanel Grid.Row="2" Margin="0,20,0,0">
                    <Button x:Name="navDashboard" Style="{StaticResource NavButton}" Content="📊 Dashboard"/>
                    <Button x:Name="navUsers" Style="{StaticResource NavButton}" Content="👤 User Management"/>
                    <Button x:Name="navLicenses" Style="{StaticResource NavButton}" Content="🔑 Licenses"/>
                    <Button x:Name="navGroups" Style="{StaticResource NavButton}" Content="👥 Groups"/>
                    <Button x:Name="navReports" Style="{StaticResource NavButton}" Content="📈 Reports"/>
                    <Button x:Name="navSettings" Style="{StaticResource NavButton}" Content="⚙️ Settings"/>
                </StackPanel>
                
                <!-- Version -->
                <TextBlock Grid.Row="3" Text="v1.0.0" FontSize="11" Foreground="#FF666666" 
                          HorizontalAlignment="Center" Margin="0,15"/>
            </Grid>
        </Border>
        
        <!-- Main Content -->
        <Grid Grid.Column="1" Background="#FF1E1E1E">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            
            <!-- Header -->
            <Border Grid.Row="0" Background="#FF252526" BorderBrush="#FF3E3E42" BorderThickness="0,0,0,1" Padding="25,20">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock x:Name="pageTitle" Grid.Column="0" Text="Dashboard" FontSize="24" FontWeight="SemiBold" Foreground="White"/>
                    <Button x:Name="btnConnect" Grid.Column="1" Style="{StaticResource ModernButton}" 
                           Content="Connect to M365" Padding="20,10"/>
                </Grid>
            </Border>
            
            <!-- Content Area -->
            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <ContentControl x:Name="contentArea" Margin="25">
                    
                    <!-- Dashboard View (Default) -->
                    <Grid x:Name="viewDashboard">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        
                        <!-- Stats Cards -->
                        <UniformGrid Grid.Row="0" Columns="4" Margin="0,0,0,20">
                            <Border Style="{StaticResource CardBorder}">
                                <StackPanel>
                                    <TextBlock Text="👤" FontSize="24" Margin="0,0,0,10"/>
                                    <TextBlock x:Name="statTotalUsers" Text="-" FontSize="32" FontWeight="Bold" Foreground="White"/>
                                    <TextBlock Text="Total Users" FontSize="12" Foreground="#FF858585" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <Border Style="{StaticResource CardBorder}">
                                <StackPanel>
                                    <TextBlock Text="✅" FontSize="24" Margin="0,0,0,10"/>
                                    <TextBlock x:Name="statActiveUsers" Text="-" FontSize="32" FontWeight="Bold" Foreground="#FF107C10"/>
                                    <TextBlock Text="Active Users" FontSize="12" Foreground="#FF858585" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <Border Style="{StaticResource CardBorder}">
                                <StackPanel>
                                    <TextBlock Text="🔑" FontSize="24" Margin="0,0,0,10"/>
                                    <TextBlock x:Name="statLicenses" Text="-" FontSize="32" FontWeight="Bold" Foreground="#FFFFA500"/>
                                    <TextBlock Text="Licenses Used" FontSize="12" Foreground="#FF858585" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                            
                            <Border Style="{StaticResource CardBorder}">
                                <StackPanel>
                                    <TextBlock Text="👥" FontSize="24" Margin="0,0,0,10"/>
                                    <TextBlock x:Name="statGroups" Text="-" FontSize="32" FontWeight="Bold" Foreground="#FF0078D4"/>
                                    <TextBlock Text="Groups" FontSize="12" Foreground="#FF858585" Margin="0,5,0,0"/>
                                </StackPanel>
                            </Border>
                        </UniformGrid>
                        
                        <!-- Quick Actions -->
                        <Border Grid.Row="1" Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Quick Actions" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,20"/>
                                <WrapPanel>
                                    <Button x:Name="btnQuickBulkImport" Style="{StaticResource ModernButton}" 
                                           Content="📊 Bulk Import Users" Margin="0,0,10,10" Width="180"/>
                                    <Button x:Name="btnQuickCreateUser" Style="{StaticResource ModernButton}" 
                                           Content="➕ Create Single User" Margin="0,0,10,10" Width="180"/>
                                    <Button x:Name="btnQuickAssignLicense" Style="{StaticResource ModernButton}" 
                                           Content="🔑 Assign Licenses" Margin="0,0,10,10" Width="180"/>
                                    <Button x:Name="btnQuickExport" Style="{StaticResource ModernButton}" 
                                           Content="📤 Export Users" Margin="0,0,10,10" Width="180"/>
                                </WrapPanel>
                            </StackPanel>
                        </Border>
                    </Grid>
                    
                </ContentControl>
            </ScrollViewer>
        </Grid>
    </Grid>
</Window>
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
