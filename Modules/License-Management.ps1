# License-Management.ps1
# License management functions

function Show-LicenseView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "License Management"
    
    if (-not (Assert-GraphConnection)) { return }
    
    [xml]$licenseXaml = @"&lt;Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"&gt;
    &lt;Grid.RowDefinitions&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="*"/&gt;
    &lt;/Grid.RowDefinitions&gt;
    
    &lt;TextBlock Grid.Row="0" Text="Available Licenses" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/&gt;
    
    &lt;Border Grid.Row="1" Style="{StaticResource CardBorder}"&gt;
        &lt;DataGrid x:Name="dgLicenses" AutoGenerateColumns="False" IsReadOnly="True"
                 GridLinesVisibility="Horizontal" Background="Transparent"
                 RowBackground="&#35;FF2D2D30" AlternatingRowBackground="&#35;FF252526"
                 Foreground="White" BorderThickness="0" HeadersVisibility="Column"&gt;
            &lt;DataGrid.Columns&gt;
                &lt;DataGridTextColumn Header="License Name" Binding="{Binding SkuPartNumber}" Width="*"/&gt;
                &lt;DataGridTextColumn Header="Total" Binding="{Binding PrepaidUnits.Enabled}" Width="100"/&gt;
                &lt;DataGridTextColumn Header="Consumed" Binding="{Binding ConsumedUnits}" Width="100"/&gt;
                &lt;DataGridTextColumn Header="Available" Binding="{Binding AvailableUnits}" Width="100"/&gt;
            &lt;/DataGrid.Columns&gt;
        &lt;/DataGrid&gt;
    &lt;/Border&gt;
&lt;/Grid&gt;
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
