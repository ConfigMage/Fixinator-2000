Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Remote GP Report Tool'
$form.Size = New-Object System.Drawing.Size(400,250)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(280,20)
$label.Text = 'Enter Computer Name:'
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,45)
$textBox.Size = New-Object System.Drawing.Size(360,20)
$form.Controls.Add($textBox)

$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(10,80)
$button.Size = New-Object System.Drawing.Size(360,30)
$button.Text = 'Generate GP Report'
$form.Controls.Add($button)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10,120)
$statusLabel.Size = New-Object System.Drawing.Size(360,80)
$statusLabel.Text = 'Ready...'
$form.Controls.Add($statusLabel)

$button.Add_Click({
    $computerName = $textBox.Text.Trim()
    
    if ([string]::IsNullOrEmpty($computerName)) {
        $statusLabel.Text = 'ERROR: Please enter a computer name.'
        $statusLabel.ForeColor = 'Red'
        return
    }
    
    $statusLabel.Text = "Connecting to $computerName..."
    $statusLabel.ForeColor = 'Blue'
    $form.Refresh()
    
    try {
        if (-not (Test-Path "C:\temp")) {
            New-Item -Path "C:\temp" -ItemType Directory -Force | Out-Null
        }
        
        $outputFile = "C:\temp\$computerName.html"
        
        $statusLabel.Text = "Generating GP Report on $computerName..."
        $form.Refresh()
        
        Invoke-Command -ComputerName $computerName -ScriptBlock {
            $tempFile = "$env:TEMP\gpreport.html"
            gpresult /h $tempFile /f
            Get-Content $tempFile -Raw
            Remove-Item $tempFile -Force
        } -OutVariable reportContent -ErrorAction Stop | Out-Null
        
        $reportContent | Out-File -FilePath $outputFile -Encoding UTF8
        
        $statusLabel.Text = "SUCCESS!`nReport saved to: $outputFile"
        $statusLabel.ForeColor = 'Green'
        
    } catch {
        $statusLabel.Text = "ERROR: $($_.Exception.Message)"
        $statusLabel.ForeColor = 'Red'
    }
})

$form.AcceptButton = $button

$form.ShowDialog()