# User Data Copy Tool
# PowerShell GUI for copying user data between computers

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to create the lifecycle directory
function New-LifecycleDirectory {
    param(
        [string]$ComputerName
    )

    try {
        $lifecyclePath = "\\$ComputerName\C$\lifecycle"

        if (-not (Test-Path $lifecyclePath)) {
            New-Item -Path $lifecyclePath -ItemType Directory -Force | Out-Null
            Write-Host "Created lifecycle directory on $ComputerName"
            return $true
        }
        return $true
    }
    catch {
        Write-Host "Error creating lifecycle directory: $_"
        return $false
    }
}

# Function to copy files with progress
function Copy-UserData {
    param(
        [string]$OldComputer,
        [string]$NewComputer,
        [string]$Username,
        [array]$FoldersToProcess,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.RichTextBox]$LogBox
    )

    $LogBox.AppendText("Starting copy process...`r`n")
    $LogBox.AppendText("From: $OldComputer`r`n")
    $LogBox.AppendText("To: $NewComputer`r`n")
    $LogBox.AppendText("User: $Username`r`n")
    $LogBox.AppendText("Folders: $($FoldersToProcess -join ', ')`r`n")
    $LogBox.AppendText("-" * 50 + "`r`n")

    # Create lifecycle directory first
    $StatusLabel.Text = "Creating lifecycle directory..."
    if (-not (New-LifecycleDirectory -ComputerName $NewComputer)) {
        $LogBox.SelectionColor = [System.Drawing.Color]::Red
        $LogBox.AppendText("ERROR: Failed to create lifecycle directory on $NewComputer`r`n")
        $StatusLabel.Text = "Error: Failed to create lifecycle directory"
        return $false
    }

    $totalFolders = $FoldersToProcess.Count
    $currentFolder = 0

    foreach ($folder in $FoldersToProcess) {
        $currentFolder++
        $progress = ($currentFolder / $totalFolders) * 100
        $ProgressBar.Value = [Math]::Min($progress, 100)

        # Determine the actual folder path based on the folder name
        switch ($folder) {
            "Documents" {
                $folderPath = "Users\$Username\Documents"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\$folder"
            }
            "Desktop" {
                $folderPath = "Users\$Username\Desktop"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\$folder"
            }
            "Pictures" {
                $folderPath = "Users\$Username\Pictures"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\$folder"
            }
            "Downloads" {
                $folderPath = "Users\$Username\Downloads"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\$folder"
            }
            "Chrome Bookmarks" {
                $folderPath = "Users\$Username\AppData\Local\Google\Chrome\User Data\Default"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\ChromeProfile"
            }
            "Edge Bookmarks" {
                $folderPath = "Users\$Username\AppData\Local\Microsoft\Edge\User Data\Default"
                $sourcePath = "\\$OldComputer\C$\$folderPath"
                $destPath = "\\$NewComputer\C$\lifecycle\Users\$Username\EdgeProfile"
            }
        }

        $StatusLabel.Text = "Copying $folder..."
        $LogBox.AppendText("Copying $folder...`r`n")
        $LogBox.AppendText("  Source: $sourcePath`r`n")
        $LogBox.AppendText("  Destination: $destPath`r`n")

        try {
            # Check if source exists
            if (Test-Path $sourcePath) {
                # Create destination directory structure
                $destParent = Split-Path $destPath -Parent
                if (-not (Test-Path $destParent)) {
                    New-Item -Path $destParent -ItemType Directory -Force | Out-Null
                }

                # Use robocopy for better performance and reliability
                $robocopyArgs = @(
                    $sourcePath,
                    $destPath,
                    "/E",           # Copy subdirectories, including empty ones
                    "/COPY:DAT",    # Copy Data, Attributes, Timestamps
                    "/R:3",         # Retry 3 times
                    "/W:5",         # Wait 5 seconds between retries
                    "/NP",          # No progress (we're using our own)
                    "/LOG+:$env:TEMP\copytool_log.txt"
                )

                $result = Start-Process -FilePath "robocopy.exe" -ArgumentList $robocopyArgs -Wait -PassThru -NoNewWindow

                if ($result.ExitCode -le 7) {  # Robocopy exit codes 0-7 indicate success
                    $LogBox.SelectionColor = [System.Drawing.Color]::Green
                    $LogBox.AppendText("  SUCCESS: $folder copied successfully`r`n")
                } else {
                    $LogBox.SelectionColor = [System.Drawing.Color]::Orange
                    $LogBox.AppendText("  WARNING: $folder copied with errors (Exit code: $($result.ExitCode))`r`n")
                }
            }
            else {
                $LogBox.SelectionColor = [System.Drawing.Color]::Orange
                $LogBox.AppendText("  WARNING: Source folder not found: $sourcePath`r`n")
            }
        }
        catch {
            $LogBox.SelectionColor = [System.Drawing.Color]::Red
            $LogBox.AppendText("  ERROR: Failed to copy $folder - $_`r`n")
        }

        $LogBox.SelectionColor = [System.Drawing.Color]::Black
    }

    $ProgressBar.Value = 100
    $StatusLabel.Text = "Copy process completed"
    $LogBox.AppendText("-" * 50 + "`r`n")
    $LogBox.AppendText("Copy process completed at $(Get-Date)`r`n")

    return $true
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "User Data Copy Tool"
$form.Size = New-Object System.Drawing.Size(600, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "User Data Migration Tool"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size = New-Object System.Drawing.Size(560, 30)
$titleLabel.Location = New-Object System.Drawing.Point(20, 10)
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($titleLabel)

# Old Computer Name
$lblOldComputer = New-Object System.Windows.Forms.Label
$lblOldComputer.Text = "Old Computer Name:"
$lblOldComputer.Location = New-Object System.Drawing.Point(20, 50)
$lblOldComputer.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($lblOldComputer)

$txtOldComputer = New-Object System.Windows.Forms.TextBox
$txtOldComputer.Location = New-Object System.Drawing.Point(180, 50)
$txtOldComputer.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($txtOldComputer)

# New Computer Name
$lblNewComputer = New-Object System.Windows.Forms.Label
$lblNewComputer.Text = "New Computer Name:"
$lblNewComputer.Location = New-Object System.Drawing.Point(20, 80)
$lblNewComputer.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($lblNewComputer)

$txtNewComputer = New-Object System.Windows.Forms.TextBox
$txtNewComputer.Location = New-Object System.Drawing.Point(180, 80)
$txtNewComputer.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($txtNewComputer)

# Username
$lblUsername = New-Object System.Windows.Forms.Label
$lblUsername.Text = "Username:"
$lblUsername.Location = New-Object System.Drawing.Point(20, 110)
$lblUsername.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($lblUsername)

$txtUsername = New-Object System.Windows.Forms.TextBox
$txtUsername.Location = New-Object System.Drawing.Point(180, 110)
$txtUsername.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($txtUsername)

# GroupBox for folder selection
$grpFolders = New-Object System.Windows.Forms.GroupBox
$grpFolders.Text = "Select Folders to Copy"
$grpFolders.Location = New-Object System.Drawing.Point(20, 140)
$grpFolders.Size = New-Object System.Drawing.Size(540, 120)
$form.Controls.Add($grpFolders)

# Checkboxes for folders
$chkDocuments = New-Object System.Windows.Forms.CheckBox
$chkDocuments.Text = "Documents"
$chkDocuments.Location = New-Object System.Drawing.Point(20, 25)
$chkDocuments.Size = New-Object System.Drawing.Size(150, 20)
$chkDocuments.Checked = $true
$grpFolders.Controls.Add($chkDocuments)

$chkDesktop = New-Object System.Windows.Forms.CheckBox
$chkDesktop.Text = "Desktop"
$chkDesktop.Location = New-Object System.Drawing.Point(200, 25)
$chkDesktop.Size = New-Object System.Drawing.Size(150, 20)
$chkDesktop.Checked = $true
$grpFolders.Controls.Add($chkDesktop)

$chkPictures = New-Object System.Windows.Forms.CheckBox
$chkPictures.Text = "Pictures"
$chkPictures.Location = New-Object System.Drawing.Point(380, 25)
$chkPictures.Size = New-Object System.Drawing.Size(150, 20)
$chkPictures.Checked = $true
$grpFolders.Controls.Add($chkPictures)

$chkDownloads = New-Object System.Windows.Forms.CheckBox
$chkDownloads.Text = "Downloads"
$chkDownloads.Location = New-Object System.Drawing.Point(20, 55)
$chkDownloads.Size = New-Object System.Drawing.Size(150, 20)
$grpFolders.Controls.Add($chkDownloads)

$chkChromeBookmarks = New-Object System.Windows.Forms.CheckBox
$chkChromeBookmarks.Text = "Chrome Bookmarks"
$chkChromeBookmarks.Location = New-Object System.Drawing.Point(200, 55)
$chkChromeBookmarks.Size = New-Object System.Drawing.Size(150, 20)
$grpFolders.Controls.Add($chkChromeBookmarks)

$chkEdgeBookmarks = New-Object System.Windows.Forms.CheckBox
$chkEdgeBookmarks.Text = "Edge Bookmarks"
$chkEdgeBookmarks.Location = New-Object System.Drawing.Point(380, 55)
$chkEdgeBookmarks.Size = New-Object System.Drawing.Size(150, 20)
$grpFolders.Controls.Add($chkEdgeBookmarks)

# Select All / Deselect All buttons
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(20, 85)
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 25)
$btnSelectAll.Add_Click({
    $chkDocuments.Checked = $true
    $chkDesktop.Checked = $true
    $chkPictures.Checked = $true
    $chkDownloads.Checked = $true
    $chkChromeBookmarks.Checked = $true
    $chkEdgeBookmarks.Checked = $true
})
$grpFolders.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Location = New-Object System.Drawing.Point(130, 85)
$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 25)
$btnDeselectAll.Add_Click({
    $chkDocuments.Checked = $false
    $chkDesktop.Checked = $false
    $chkPictures.Checked = $false
    $chkDownloads.Checked = $false
    $chkChromeBookmarks.Checked = $false
    $chkEdgeBookmarks.Checked = $false
})
$grpFolders.Controls.Add($btnDeselectAll)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 270)
$progressBar.Size = New-Object System.Drawing.Size(540, 25)
$progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progressBar)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready"
$statusLabel.Location = New-Object System.Drawing.Point(20, 305)
$statusLabel.Size = New-Object System.Drawing.Size(540, 20)
$form.Controls.Add($statusLabel)

# Log/Output Box
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Operation Log:"
$lblLog.Location = New-Object System.Drawing.Point(20, 335)
$lblLog.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblLog)

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Location = New-Object System.Drawing.Point(20, 360)
$logBox.Size = New-Object System.Drawing.Size(540, 200)
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

# Buttons
$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "Start Copy"
$btnCopy.Location = New-Object System.Drawing.Point(150, 580)
$btnCopy.Size = New-Object System.Drawing.Size(120, 40)
$btnCopy.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$btnCopy.BackColor = [System.Drawing.Color]::LightGreen
$btnCopy.Add_Click({
    # Validate inputs
    if ([string]::IsNullOrWhiteSpace($txtOldComputer.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the old computer name.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    if ([string]::IsNullOrWhiteSpace($txtNewComputer.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the new computer name.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    if ([string]::IsNullOrWhiteSpace($txtUsername.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the username.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Get selected folders
    $selectedFolders = @()
    if ($chkDocuments.Checked) { $selectedFolders += "Documents" }
    if ($chkDesktop.Checked) { $selectedFolders += "Desktop" }
    if ($chkPictures.Checked) { $selectedFolders += "Pictures" }
    if ($chkDownloads.Checked) { $selectedFolders += "Downloads" }
    if ($chkChromeBookmarks.Checked) { $selectedFolders += "Chrome Bookmarks" }
    if ($chkEdgeBookmarks.Checked) { $selectedFolders += "Edge Bookmarks" }

    if ($selectedFolders.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one folder to copy.", "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Confirm action
    $message = "This will copy the following folders for user '$($txtUsername.Text)':`n`n"
    $message += "$($selectedFolders -join ', ')`n`n"
    $message += "From: \\$($txtOldComputer.Text)\C$\Users\$($txtUsername.Text)\`n"
    $message += "To: \\$($txtNewComputer.Text)\C$\lifecycle\Users\$($txtUsername.Text)\`n`n"
    $message += "Do you want to proceed?"

    $result = [System.Windows.Forms.MessageBox]::Show($message, "Confirm Copy", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Disable controls during copy
        $btnCopy.Enabled = $false
        $btnCancel.Enabled = $false
        $txtOldComputer.Enabled = $false
        $txtNewComputer.Enabled = $false
        $txtUsername.Enabled = $false
        $grpFolders.Enabled = $false

        # Clear log
        $logBox.Clear()
        $progressBar.Value = 0

        # Start copy process
        Copy-UserData -OldComputer $txtOldComputer.Text `
                     -NewComputer $txtNewComputer.Text `
                     -Username $txtUsername.Text `
                     -FoldersToProcess $selectedFolders `
                     -ProgressBar $progressBar `
                     -StatusLabel $statusLabel `
                     -LogBox $logBox

        # Re-enable controls
        $btnCopy.Enabled = $true
        $btnCancel.Enabled = $true
        $txtOldComputer.Enabled = $true
        $txtNewComputer.Enabled = $true
        $txtUsername.Enabled = $true
        $grpFolders.Enabled = $true

        [System.Windows.Forms.MessageBox]::Show("Copy process completed. Check the log for details.", "Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})
$form.Controls.Add($btnCopy)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear Form"
$btnClear.Location = New-Object System.Drawing.Point(280, 580)
$btnClear.Size = New-Object System.Drawing.Size(100, 40)
$btnClear.Add_Click({
    $txtOldComputer.Clear()
    $txtNewComputer.Clear()
    $txtUsername.Clear()
    $logBox.Clear()
    $progressBar.Value = 0
    $statusLabel.Text = "Ready"
    $chkDocuments.Checked = $true
    $chkDesktop.Checked = $true
    $chkPictures.Checked = $true
    $chkDownloads.Checked = $false
    $chkChromeBookmarks.Checked = $false
    $chkEdgeBookmarks.Checked = $false
})
$form.Controls.Add($btnClear)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Exit"
$btnCancel.Location = New-Object System.Drawing.Point(390, 580)
$btnCancel.Size = New-Object System.Drawing.Size(100, 40)
$btnCancel.Add_Click({ $form.Close() })
$form.Controls.Add($btnCancel)

# Show the form
$form.ShowDialog() | Out-Null