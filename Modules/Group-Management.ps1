# Group-Management.ps1
# Group management functions

function Show-GroupView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "Group Management"
    
    if (-not (Assert-GraphConnection)) { return }
    
    [xml]$groupXaml = @"
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    
    <TextBlock Grid.Row="0" Text="Azure AD Groups" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15" Foreground="White"/>
    
    <Border Grid.Row="1" Style="{StaticResource CardBorder}">
        <DataGrid x:Name="dgGroups" AutoGenerateColumns="False" IsReadOnly="True"
                 GridLinesVisibility="Horizontal" Background="Transparent"
                 RowBackground="#FF2D2D30" AlternatingRowBackground="#FF252526"
                 Foreground="White" BorderThickness="0" HeadersVisibility="Column">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Group Name" Binding="{Binding DisplayName}" Width="*"/>
                <DataGridTextColumn Header="Type" Binding="{Binding GroupTypes}" Width="150"/>
                <DataGridTextColumn Header="Members" Binding="{Binding MemberCount}" Width="100"/>
                <DataGridTextColumn Header="Visibility" Binding="{Binding Visibility}" Width="100"/>
            </DataGrid.Columns>
        </DataGrid>
    </Border>
</Grid>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $groupXaml
    $view = [Windows.Markup.XamlReader]::Load($reader)
    
    # Load group data
    try {
        $groups = Get-MgGroup -All | Select-Object DisplayName, GroupTypes, Visibility,
            @{N='MemberCount'; E={ (Get-MgGroupMember -GroupId $_.Id).Count }}
        $view.FindName("dgGroups").ItemsSource = $groups
    }
    catch {
        Write-Warning "Failed to load groups: $_"
    }
    
    Switch-View -Ui $Ui -View $view
}
