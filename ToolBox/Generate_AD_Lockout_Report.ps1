Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'AD Locked Out Accounts'
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,10)
$label.Size = New-Object System.Drawing.Size(500,20)
$label.Text = 'Currently Locked Out Active Directory Accounts:'
$form.Controls.Add($label)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10,530)
$statusLabel.Size = New-Object System.Drawing.Size(600,20)
$statusLabel.Text = 'Ready'
$form.Controls.Add($statusLabel)

$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(10,40)
$dataGridView.Size = New-Object System.Drawing.Size(860,440)
$dataGridView.AutoSizeColumnsMode = 'Fill'
$dataGridView.AllowUserToAddRows = $false
$dataGridView.ReadOnly = $true
$dataGridView.SelectionMode = 'FullRowSelect'
$form.Controls.Add($dataGridView)

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = New-Object System.Drawing.Point(10,490)
$refreshButton.Size = New-Object System.Drawing.Size(100,30)
$refreshButton.Text = 'Refresh'
$form.Controls.Add($refreshButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Location = New-Object System.Drawing.Point(120,490)
$exportButton.Size = New-Object System.Drawing.Size(100,30)
$exportButton.Text = 'Export to CSV'
$form.Controls.Add($exportButton)

function Get-LockedOutAccounts {
    $statusLabel.Text = 'Querying Active Directory...'
    $statusLabel.Refresh()
    
    try {
        $lockedAccounts = Search-ADAccount -LockedOut -UsersOnly | Select-Object Name, SamAccountName, DistinguishedName
        
        if ($lockedAccounts) {
            $accountDetails = @()
            
            foreach ($account in $lockedAccounts) {
                try {
                    $user = Get-ADUser $account.SamAccountName -Properties LockoutTime, BadPasswordTime, LockedOut, LastBadPasswordAttempt
                    
                    $lockoutTime = if ($user.LockoutTime -ne 0) {
                        [DateTime]::FromFileTime($user.LockoutTime)
                    } else {
                        "N/A"
                    }
                    
                    $domain = (Get-ADDomain).PDCEmulator
                    
                    $lockingDC = "N/A"
                    try {
                        $events = Get-WinEvent -ComputerName $domain -FilterHashtable @{
                            LogName = 'Security'
                            Id = 4740
                        } -MaxEvents 100 -ErrorAction SilentlyContinue
                        
                        $userEvent = $events | Where-Object {
                            $_.Properties[0].Value -eq $account.SamAccountName
                        } | Select-Object -First 1
                        
                        if ($userEvent) {
                            $lockingDC = $userEvent.Properties[1].Value
                        }
                    } catch {
                        $lockingDC = $domain
                    }
                    
                    $accountDetails += [PSCustomObject]@{
                        'SAM Account Name' = $account.SamAccountName
                        'Display Name' = $account.Name
                        'Lockout Time' = $lockoutTime
                        'Domain Controller' = $lockingDC
                        'Distinguished Name' = $account.DistinguishedName
                    }
                } catch {
                    Write-Warning "Error processing $($account.SamAccountName): $_"
                }
            }
            
            $dataGridView.DataSource = [System.Collections.ArrayList]@($accountDetails)
            $statusLabel.Text = "Found $($accountDetails.Count) locked out account(s)"
        } else {
            $dataGridView.DataSource = $null
            $statusLabel.Text = 'No locked out accounts found'
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error querying Active Directory: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $statusLabel.Text = 'Error occurred'
    }
}

$refreshButton.Add_Click({
    Get-LockedOutAccounts
})

$exportButton.Add_Click({
    if ($dataGridView.Rows.Count -gt 0) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "LockedAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $data = @()
                foreach ($row in $dataGridView.Rows) {
                    $data += [PSCustomObject]@{
                        'SAM Account Name' = $row.Cells['SAM Account Name'].Value
                        'Display Name' = $row.Cells['Display Name'].Value
                        'Lockout Time' = $row.Cells['Lockout Time'].Value
                        'Domain Controller' = $row.Cells['Domain Controller'].Value
                        'Distinguished Name' = $row.Cells['Distinguished Name'].Value
                    }
                }
                $data | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Data exported successfully to $($saveDialog.FileName)", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting data: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
})

Get-LockedOutAccounts

[void]$form.ShowDialog()