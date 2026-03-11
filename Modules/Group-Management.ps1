# Group-Management.ps1
# Group management functions

function Show-GroupView {
    param([Parameter(Mandatory=$true)] $Ui)
    Set-PageTitle -Ui $Ui -Title "Group Management"
    
    if (-not (Assert-GraphConnection)) { return }
    
    [xml]$groupXaml = @"&lt;Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"&gt;
    &lt;Grid.RowDefinitions&gt;
        &lt;RowDefinition Height="Auto"/&gt;
        &lt;RowDefinition Height="*"/&gt;
    &lt;/Grid.RowDefinitions&gt;
    
    &lt;TextBlock Grid.Row="0" Text="Azure AD Groups" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,15"/&gt;
    
    &lt;Border Grid.Row="1" Style="{StaticResource CardBorder}"&gt;
        &lt;DataGrid x:Name="dgGroups" AutoGenerateColumns="False" IsReadOnly="True"
                 GridLinesVisibility="Horizontal" Background="Transparent"
                 RowBackground="&#35;FF2D2D30" AlternatingRowBackground="&#35;FF252526"
                 Foreground="White" BorderThickness="0" HeadersVisibility="Column"&gt;
            &lt;DataGrid.Columns&gt;
                &lt;DataGridTextColumn Header="Group Name" Binding="{Binding DisplayName}" Width="*"/&gt;
                &lt;DataGridTextColumn Header="Type" Binding="{Binding GroupTypes}" Width="150"/&gt;
                &lt;DataGridTextColumn Header="Members" Binding="{Binding MemberCount}" Width="100"/&gt;
                &lt;DataGridTextColumn Header="Visibility" Binding="{Binding Visibility}" Width="100"/&gt;
            &lt;/DataGrid.Columns&gt;
        &lt;/DataGrid&gt;
    &lt;/Border&gt;
&lt;/Grid&gt;
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
