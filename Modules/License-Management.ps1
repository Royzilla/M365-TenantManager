# License-Management.ps1
# License management functions

function Show-LicenseView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "License Management"
    
    if (-not (Assert-GraphConnection)) { return }
    
    [xml]$licenseXaml = @"
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    
    <TextBlock Grid.Row="0" Text="Available Licenses" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15" Foreground="White"/>
    
    <Border Grid.Row="1" Style="{StaticResource CardBorder}">
        <DataGrid x:Name="dgLicenses" AutoGenerateColumns="False" IsReadOnly="True"
                 GridLinesVisibility="Horizontal" Background="Transparent"
                 RowBackground="#FF2D2D30" AlternatingRowBackground="#FF252526"
                 Foreground="White" BorderThickness="0" HeadersVisibility="Column">
            <DataGrid.Columns>
                <DataGridTextColumn Header="License Name" Binding="{Binding SkuPartNumber}" Width="*"/>
                <DataGridTextColumn Header="Total" Binding="{Binding PrepaidUnits.Enabled}" Width="100"/>
                <DataGridTextColumn Header="Consumed" Binding="{Binding ConsumedUnits}" Width="100"/>
                <DataGridTextColumn Header="Available" Binding="{Binding AvailableUnits}" Width="100"/>
            </DataGrid.Columns>
        </DataGrid>
    </Border>
</Grid>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $licenseXaml
    $view = [Windows.Markup.XamlReader]::Load($reader)
    
    # Load license data
    try {
        $skus = Get-MgSubscribedSku | Select-Object SkuPartNumber, ConsumedUnits, 
            @{N='PrepaidUnits'; E={$_.PrepaidUnits}},
            @{N='AvailableUnits'; E={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}
        $view.FindName("dgLicenses").ItemsSource = $skus
    }
    catch {
        Write-Warning "Failed to load licenses: $_"
    }
    
    Switch-View -Ui $Ui -View $view
}
