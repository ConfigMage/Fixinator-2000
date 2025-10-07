Add-Type -AssemblyName PresentationCore, PresentationFramework
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Printer Tool" Width="820" Height="560" WindowStartupLocation="CenterScreen">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>

    <Label Grid.Row="0" Grid.Column="0" Content="Target computer(FQDN):" VerticalAlignment="Center"/>
    <TextBox Grid.Row="0" Grid.Column="1" Name="txt_target" MinWidth="280" Margin="6,0,6,0"/>
    <StackPanel Grid.Row="0" Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button Name="btn_listprinters" Content="List Printers" Width="120" Margin="0,0,8,0"/>
      <Button Name="btn_exportcsv" Content="Export CSV" Width="120"/>
    </StackPanel>

    <ListBox Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="3" Name="lst_printers" SelectionMode="Extended" Margin="0,10,0,10"/>

    <StackPanel Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Orientation="Horizontal" Margin="0,0,0,10">
      <CheckBox Name="chk_removeports" Content="Also remove printer ports" IsChecked="True" VerticalAlignment="Center"/>
      <Button Name="btn_removeprinters" Content="Remove Selected" Width="150" Margin="12,0,0,0"/>
    </StackPanel>

    <TextBlock Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Text="Enter the printer name. Example ITS-RICIM5000" FontSize="14" Margin="0,0,0,6"/>

    <Grid Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="3" Margin="0,0,0,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Label Grid.Column="0" Content="Printer name:" VerticalAlignment="Center"/>
      <TextBox Grid.Column="1" Name="txt_printername" Margin="6,0,6,0"/>
      <Button Grid.Column="2" Name="btn_addprinter" Content="Add Printer" Width="150"/>
    </Grid>

    <TextBlock Grid.Row="5" Grid.Column="0" Grid.ColumnSpan="3" Text="Enter the printer name that has faxing enabled. Example ITS-RICIM5000" FontSize="14" Margin="0,0,0,6"/>

    <Grid Grid.Row="6" Grid.Column="0" Grid.ColumnSpan="3">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Label Grid.Column="0" Content="Fax name:" VerticalAlignment="Center"/>
      <TextBox Grid.Column="1" Name="txt_faxname" Margin="6,0,6,0"/>
      <Button Grid.Column="2" Name="btn_addfaxes" Content="Add Fax" Width="150"/>
    </Grid>
  </Grid>
</Window>
"@

Set-StrictMode -Version Latest

Function Get-Target {
  if ($txt_target -and -not [string]::IsNullOrWhiteSpace($txt_target.Text)) { return $txt_target.Text.Trim() }
  return $null
}

Function Get-Remote-Printers {
  param([string] $Computer, [pscredential] $Credential)
  $sb = { Get-Printer | Select-Object Name, DriverName, PortName, Shared, Type }
  if ($Computer) {
    if ($Credential) { return Invoke-Command -ComputerName $Computer -Credential $Credential -ScriptBlock $sb -ErrorAction Stop }
    return Invoke-Command -ComputerName $Computer -ScriptBlock $sb -ErrorAction Stop
  }
  return & $sb
}

Function List-Printers {
  $target = Get-Target
  $cred = $null
  try {
    $printers = Get-Remote-Printers -Computer $target -Credential $cred
    $lst_printers.Items.Clear()
    foreach ($p in $printers) {
      $display = "{0} | {1} | Port:{2}" -f $p.Name, $p.DriverName, $p.PortName
      $obj = [pscustomobject]@{
        Display=$display
        Name=$p.Name
        DriverName=$p.DriverName
        PortName=$p.PortName
        Shared=$p.Shared
        Type=$p.Type
        Computer=($(if ($target) { $target } else { 'local' }))
      }
      [void]$lst_printers.Items.Add($obj)
    }
    $lst_printers.DisplayMemberPath = 'Display'
    if (-not $printers) { [System.Windows.MessageBox]::Show('No printers found on target.', 'Info') | Out-Null }
  } catch { [System.Windows.MessageBox]::Show("Failed to list printers: " + $_.Exception.Message, 'Error') | Out-Null }
}

Function Remove-Remote-Printers {
  param([string] $Computer, [string[]] $PrinterNames, [string[]] $PortNames, [switch] $RemovePorts, [pscredential] $Credential)
  $sb = {
    param($Names, $Ports, $RemovePorts)
    $ErrorActionPreference = 'Stop'
    $results = @()
    for ($i=0; $i -lt $Names.Count; $i++) {
      $n = $Names[$i]
      $port = $null; if ($Ports -and $i -lt $Ports.Count) { $port = $Ports[$i] }
      try {
        if (Get-Printer -Name $n -ErrorAction SilentlyContinue) { Remove-Printer -Name $n -ErrorAction Stop }
        if ($RemovePorts -and $port) { if (Get-PrinterPort -Name $port -ErrorAction SilentlyContinue) { Remove-PrinterPort -Name $port -ErrorAction Stop } }
        $results += [pscustomobject]@{ Name=$n; Status='Removed' }
      } catch { $results += [pscustomobject]@{ Name=$n; Status=("Failed: " + $_.Exception.Message) } }
    }
    return $results
  }
  if (-not $Computer) { return & $sb -ArgumentList $PrinterNames, $PortNames, [bool]$RemovePorts }
  if ($Credential) { return Invoke-Command -ComputerName $Computer -ScriptBlock $sb -ArgumentList $PrinterNames, $PortNames, [bool]$RemovePorts -Credential $Credential }
  return Invoke-Command -ComputerName $Computer -ScriptBlock $sb -ArgumentList $PrinterNames, $PortNames, [bool]$RemovePorts
}

Function Remove-Selected-Printers {
  if (-not $lst_printers.SelectedItems -or $lst_printers.SelectedItems.Count -eq 0) { [System.Windows.MessageBox]::Show('Select one or more printers to remove.', 'Selection Required') | Out-Null; return }
  $target = Get-Target
  $removePorts = $false; if ($chk_removeports -and $chk_removeports.IsChecked) { $removePorts = $true }
  $names = @(); $ports = @(); foreach ($item in $lst_printers.SelectedItems) { $names += $item.Name; $ports += $item.PortName }
  $cred = $null
  if ($target) { $cred = Get-Credential -Message "Credentials for remote removal on $target" }
  try {
    $res = Remove-Remote-Printers -Computer $target -PrinterNames $names -PortNames $ports -RemovePorts:($removePorts) -Credential $cred
    $msg = ($res | ForEach-Object { "{0} - {1}" -f $_.Name, $_.Status }) -join "`n"
    [System.Windows.MessageBox]::Show($msg, 'Removal Results') | Out-Null
    List-Printers
  } catch { [System.Windows.MessageBox]::Show("Failed to remove printers: " + $_.Exception.Message, 'Error') | Out-Null }
}

Function Add-Local-Printer {
  param([string] $PrinterName, [string] $DriverName, [string] $PortHostName)
  try {
    $portName = if ($PSBoundParameters.ContainsKey('PortHostName') -and $PortHostName) { $PortHostName } else { $PrinterName }
    if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $portName -PrinterHostAddress $portName | Out-Null }
    if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) { Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $portName | Out-Null }
    return $true
  } catch { return $_.Exception.Message }
}

Function Add-Remote-Printer {
  param([string] $Computer, [string] $PrinterName, [string] $DriverName, [pscredential] $Credential, [string] $PortHostName)
  if (-not $Computer) { return Add-Local-Printer -PrinterName $PrinterName -DriverName $DriverName -PortHostName $PortHostName }
  $script = {
    param($PrinterName, $DriverName, $PortHostName)
    $ErrorActionPreference = 'Stop'
    $portName = if ($PSBoundParameters.ContainsKey('PortHostName') -and $PortHostName) { $PortHostName } else { $PrinterName }
    if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $portName -PrinterHostAddress $portName | Out-Null }
    if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) { Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $portName | Out-Null }
    'OK'
  }
  try {
    if ($Credential) { Invoke-Command -ComputerName $Computer -ScriptBlock $script -ArgumentList $PrinterName, $DriverName, $PortHostName -Credential $Credential -ErrorAction Stop | Out-Null }
    else { Invoke-Command -ComputerName $Computer -ScriptBlock $script -ArgumentList $PrinterName, $DriverName, $PortHostName -ErrorAction Stop | Out-Null }
    return $true
  } catch { return $_.Exception.Message }
}

Function Adding-Printer {
  $p = $txt_printername.Text
  if ([string]::IsNullOrWhiteSpace($p)) { [System.Windows.MessageBox]::Show('Enter a printer name.', 'Input Required') | Out-Null; return }
  $target = Get-Target
  $cred = $null
  $res = Add-Remote-Printer -Computer $target -PrinterName $p -DriverName 'RICOH PCL6 UniversalDriver V4.26' -Credential $cred
  if ($res -eq $true) { [System.Windows.MessageBox]::Show(("Added/verified '{0}' on {1}" -f $p, $(if ($target) { $target } else { 'local' })), 'Success') | Out-Null }
  else { [System.Windows.MessageBox]::Show("Install failed: $res", 'Error') | Out-Null }
}

Function Adding-Fax {
  $f = $txt_faxname.Text
  if ([string]::IsNullOrWhiteSpace($f)) { [System.Windows.MessageBox]::Show('Enter a fax base name.', 'Input Required') | Out-Null; return }
  $target = Get-Target
  $cred = $null
  $faxPrinter = "$f (Fax)"
  $res = Add-Remote-Printer -Computer $target -PrinterName $faxPrinter -DriverName 'LAN-Fax Generic' -Credential $cred -PortHostName $f
  if ($res -eq $true) { [System.Windows.MessageBox]::Show(("Added/verified '{0}' on {1}" -f $faxPrinter, $(if ($target) { $target } else { 'local' })), 'Success') | Out-Null }
  else { [System.Windows.MessageBox]::Show("Fax install failed: $res", 'Error') | Out-Null }
}

Function Export-PrintersCsv {
  if (-not $lst_printers.Items -or $lst_printers.Items.Count -eq 0) { [System.Windows.MessageBox]::Show('No printers listed to export. Click "List Printers" first.', 'Nothing to Export') | Out-Null; return }
  Add-Type -AssemblyName PresentationFramework
  $dlg = New-Object Microsoft.Win32.SaveFileDialog
  $dlg.Title = 'Export printers to CSV'
  $dlg.DefaultExt = '.csv'
  $dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
  $ok = $dlg.ShowDialog()
  if (-not $ok) { return }
  $path = $dlg.FileName
  try {
    $data = foreach ($item in $lst_printers.Items) {
      [pscustomobject]@{
        Computer   = $item.Computer
        Name       = $item.Name
        DriverName = $item.DriverName
        PortName   = $item.PortName
        Shared     = $item.Shared
        Type       = $item.Type
      }
    }
    $data | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    [System.Windows.MessageBox]::Show(("Exported {0} printers to '{1}'" -f $data.Count, $path), 'Export Complete') | Out-Null
  } catch {
    [System.Windows.MessageBox]::Show("Failed to export CSV: " + $_.Exception.Message, 'Error') | Out-Null
  }
}

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)
[xml]$xml = $Xaml
$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }
$btn_listprinters.Add_Click({ List-Printers })
$btn_removeprinters.Add_Click({ Remove-Selected-Printers })
$btn_addprinter.Add_Click({ Adding-Printer })
$btn_addfaxes.Add_Click({ Adding-Fax })
$btn_exportcsv.Add_Click({ Export-PrintersCsv })
$Window.ShowDialog()