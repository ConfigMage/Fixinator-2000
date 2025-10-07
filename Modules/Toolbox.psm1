# Toolbox.psm1
# Handles Toolbox panel and functionality - executes PowerShell scripts from ToolBox directory locally

$script:ToolboxPanel = $null
$script:ToolButtons = @()
$script:ScriptRoot = ""
$script:MainStatusLabel = $null

function Initialize-ToolboxPanel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Panel]$Panel,

        [Parameter(Mandatory = $true)]
        [string]$ScriptRoot,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel,

        [Parameter(Mandatory = $false)]
        [System.Windows.Forms.ToolStripStatusLabel]$ConnectionLabel
    )

    $script:ToolboxPanel = $Panel
    $script:ScriptRoot = $ScriptRoot
    $script:MainStatusLabel = $StatusLabel

    # Define colors for this module
    $colorPrimary = [System.Drawing.Color]::FromArgb(41, 128, 185)
    $colorTextDark = [System.Drawing.Color]::FromArgb(44, 62, 80)
    $colorTextLight = [System.Drawing.Color]::FromArgb(149, 165, 166)
    $colorBackground = [System.Drawing.Color]::FromArgb(236, 240, 241)

    # Header panel with title and description
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $headerPanel.Height = 70
    $headerPanel.BackColor = [System.Drawing.Color]::White
    $headerPanel.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

    # Title label
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(15, 12)
    $lblTitle.Size = New-Object System.Drawing.Size(500, 30)
    $lblTitle.Text = "Toolbox"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = $colorPrimary
    $headerPanel.Controls.Add($lblTitle)

    # Description label
    $lblDescription = New-Object System.Windows.Forms.Label
    $lblDescription.Location = New-Object System.Drawing.Point(15, 42)
    $lblDescription.Size = New-Object System.Drawing.Size(700, 20)
    $lblDescription.Text = "Execute local PowerShell utilities and tools"
    $lblDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblDescription.ForeColor = $colorTextLight
    $headerPanel.Controls.Add($lblDescription)

    # Add a separator line
    $separator = New-Object System.Windows.Forms.Panel
    $separator.Height = 2
    $separator.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $separator.BackColor = $colorBackground
    $headerPanel.Controls.Add($separator)

    $Panel.Controls.Add($headerPanel)

    # Tools panel
    $toolsPanel = New-Object System.Windows.Forms.Panel
    $toolsPanel.Location = New-Object System.Drawing.Point(0, 72)
    $toolsPanel.Size = New-Object System.Drawing.Size($Panel.Width, ($Panel.Height - 72))
    $toolsPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor
                           [System.Windows.Forms.AnchorStyles]::Left -bor
                           [System.Windows.Forms.AnchorStyles]::Right
    $toolsPanel.AutoScroll = $true
    $toolsPanel.BackColor = [System.Drawing.Color]::White
    $toolsPanel.Padding = New-Object System.Windows.Forms.Padding(10)

    # Get PowerShell scripts from ToolBox directory
    $toolboxPath = Join-Path $script:ScriptRoot "ToolBox"
    $toolScripts = @()

    if (Test-Path $toolboxPath) {
        $toolScripts = Get-ChildItem -Path $toolboxPath -Filter "*.ps1" | Select-Object -First 20
    }

    # Create tool buttons based on available scripts
    $buttonWidth = 240
    $buttonHeight = 120
    $buttonSpacing = 18
    $startX = 20
    $startY = 20

    $buttonIndex = 0
    foreach ($toolScript in $toolScripts) {
        $button = New-Object System.Windows.Forms.Button

        # Calculate position (3 columns layout for better use of space)
        $col = $buttonIndex % 3
        $row = [Math]::Floor($buttonIndex / 3)
        $x = $startX + ($col * ($buttonWidth + $buttonSpacing))
        $y = $startY + ($row * ($buttonHeight + $buttonSpacing))

        $button.Location = New-Object System.Drawing.Point($x, $y)
        $button.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)

        # Format the button text from the script filename
        $buttonText = $toolScript.BaseName -replace '_', ' ' -replace '-', ' '

        # Wrap long text
        if ($buttonText.Length -gt 20) {
            $words = $buttonText -split ' '
            $lines = @()
            $currentLine = ""

            foreach ($word in $words) {
                if (($currentLine + " " + $word).Length -le 18) {
                    if ($currentLine -eq "") {
                        $currentLine = $word
                    } else {
                        $currentLine += " " + $word
                    }
                } else {
                    if ($currentLine -ne "") {
                        $lines += $currentLine
                    }
                    $currentLine = $word
                }
            }
            if ($currentLine -ne "") {
                $lines += $currentLine
            }
            $button.Text = $lines -join "`n"
        } else {
            $button.Text = $buttonText
        }

        $button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 0
        $button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $button.UseCompatibleTextRendering = $false
        $button.AutoEllipsis = $true
        $button.Enabled = $true  # Buttons are enabled by default for local execution
        $button.Name = "ToolButton_$buttonIndex"
        $button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $button.TabStop = $false

        # Store the script path in the Tag property
        $button.Tag = @{
            ScriptPath = $toolScript.FullName
            ScriptName = $toolScript.BaseName
        }

        # Set button color based on script name patterns
        switch -Regex ($toolScript.BaseName) {
            'AD|Active|Lockout' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)  # Steel Blue
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 160, 210)
            }
            'Defender|Security|Toolkit' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)  # Crimson
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250, 50, 90)
            }
            'GPO|Policy|Report' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(46, 139, 87)  # Sea Green
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(76, 169, 117)
            }
            'Print' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)  # Dark Orange
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 170, 30)
            }
            'User|Data|Copy' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(75, 0, 130)  # Indigo
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(105, 30, 160)
            }
            default {
                $button.BackColor = [System.Drawing.Color]::FromArgb(105, 105, 105)  # Dim Gray
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(135, 135, 135)
            }
        }

        $button.ForeColor = [System.Drawing.Color]::White

        # Add click handler
        $button.Add_Click({
            param($sender, $eventArgs)
            Invoke-ToolboxScript -ToolButton $sender
        })

        # Add tooltip with script filename and description
        $tooltip = New-Object System.Windows.Forms.ToolTip
        $tooltip.SetToolTip($button, "Execute: $($toolScript.Name)`nClick to run this PowerShell script locally")

        # Add button to panel
        $toolsPanel.Controls.Add($button)
        $button.BringToFront()  # Ensure button is on top

        $script:ToolButtons += $button

        $buttonIndex++
    }

    # If no scripts found, show a message
    if ($toolScripts.Count -eq 0) {
        $lblNoScripts = New-Object System.Windows.Forms.Label
        $lblNoScripts.Location = New-Object System.Drawing.Point(15, 15)
        $lblNoScripts.Size = New-Object System.Drawing.Size(600, 50)
        $lblNoScripts.Text = "No PowerShell scripts found in the ToolBox directory.`nPlace your .ps1 scripts in: $toolboxPath"
        $lblNoScripts.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $lblNoScripts.ForeColor = [System.Drawing.Color]::DarkGray
        $toolsPanel.Controls.Add($lblNoScripts)
    }

    $Panel.Controls.Add($toolsPanel)
}

function Invoke-ToolboxScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Button]$ToolButton
    )

    $toolInfo = $ToolButton.Tag
    if ($null -eq $toolInfo) {
        Write-Warning "No tool information found for this button"
        return
    }

    $scriptPath = $toolInfo.ScriptPath
    $scriptName = $toolInfo.ScriptName

    if (!(Test-Path $scriptPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Script not found: $scriptPath",
            "Script Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    $script:MainStatusLabel.Text = "Executing: $scriptName..."
    $ToolButton.Enabled = $false

    try {
        Write-Log "Starting local tool: $scriptName" -Level INFO

        # Check if we have cached credentials that the script might need
        $credential = $null
        $getCredCommand = Get-Command Get-RemoteCredential -ErrorAction SilentlyContinue
        if ($getCredCommand) {
            try {
                $credential = Get-RemoteCredential
            }
            catch {
                Write-Verbose "Could not retrieve cached credentials: $_"
            }
        }

        # Execute the script locally
        # Scripts can optionally use the $credential variable if it exists
        if ($credential) {
            $result = & $scriptPath -Credential $credential -ErrorAction Stop
        }
        else {
            $result = & $scriptPath -ErrorAction Stop
        }

        $script:MainStatusLabel.Text = "$scriptName completed successfully"
        Write-Log "Tool completed: $scriptName" -Level SUCCESS

        # Display results if available
        if ($result) {
            Show-ToolResults -Title $scriptName -Results $result
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        $script:MainStatusLabel.Text = "Tool failed: $scriptName"
        Write-Log "Tool failed: $_" -Level ERROR

        [System.Windows.Forms.MessageBox]::Show(
            "Failed to execute tool '$scriptName':`n`n$errorMessage",
            "Tool Execution Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    finally {
        if ($ToolButton) {
            $ToolButton.Enabled = $true
        }
    }
}

function Show-ToolResults {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [object]$Results
    )

    # Create results form
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "$Title - Results"
    $resultsForm.Size = New-Object System.Drawing.Size(900, 700)
    $resultsForm.StartPosition = "CenterParent"
    $resultsForm.Icon = [System.Drawing.SystemIcons]::Information

    # Results text box
    $txtResults = New-Object System.Windows.Forms.TextBox
    $txtResults.Multiline = $true
    $txtResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txtResults.ReadOnly = $true
    $txtResults.Font = New-Object System.Drawing.Font("Consolas", 10)
    $txtResults.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtResults.BackColor = [System.Drawing.Color]::White

    # Format results
    if ($Results -is [Array] -or $Results -is [System.Collections.IEnumerable]) {
        $txtResults.Text = $Results | Format-Table -AutoSize | Out-String -Width 200
    }
    else {
        $txtResults.Text = $Results | Format-List | Out-String -Width 200
    }

    $resultsForm.Controls.Add($txtResults)

    # Button panel
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Height = 40
    $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom

    # Copy to clipboard button
    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = "Copy to Clipboard"
    $btnCopy.Location = New-Object System.Drawing.Point(10, 5)
    $btnCopy.Size = New-Object System.Drawing.Size(120, 30)
    $btnCopy.Add_Click({
        [System.Windows.Forms.Clipboard]::SetText($txtResults.Text)
        [System.Windows.Forms.MessageBox]::Show(
            "Results copied to clipboard!",
            "Copied",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    })
    $buttonPanel.Controls.Add($btnCopy)

    # Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(($resultsForm.Width - 110), 5)
    $btnClose.Size = New-Object System.Drawing.Size(80, 30)
    $btnClose.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $btnClose.Add_Click({
        $resultsForm.Close()
    })
    $buttonPanel.Controls.Add($btnClose)

    $resultsForm.Controls.Add($buttonPanel)
    $resultsForm.ShowDialog()
}

function Update-ToolboxButtons {
    [CmdletBinding()]
    param()

    # This function can be called to refresh the toolbox buttons
    # For example, if new scripts are added to the ToolBox directory

    if ($null -eq $script:ToolboxPanel) {
        Write-Warning "Toolbox panel not initialized"
        return
    }

    $toolboxPath = Join-Path $script:ScriptRoot "ToolBox"
    $toolScripts = @()

    if (Test-Path $toolboxPath) {
        $toolScripts = Get-ChildItem -Path $toolboxPath -Filter "*.ps1" | Select-Object -First 20
    }

    # Update existing buttons with new scripts
    for ($i = 0; $i -lt [Math]::Min($toolScripts.Count, $script:ToolButtons.Count); $i++) {
        $button = $script:ToolButtons[$i]
        $toolScript = $toolScripts[$i]

        $buttonText = $toolScript.BaseName -replace '_', ' ' -replace '-', ' '

        # Handle long text
        if ($buttonText.Length -gt 20) {
            $words = $buttonText -split ' '
            $lines = @()
            $currentLine = ""

            foreach ($word in $words) {
                if (($currentLine + " " + $word).Length -le 18) {
                    if ($currentLine -eq "") {
                        $currentLine = $word
                    } else {
                        $currentLine += " " + $word
                    }
                } else {
                    if ($currentLine -ne "") {
                        $lines += $currentLine
                    }
                    $currentLine = $word
                }
            }
            if ($currentLine -ne "") {
                $lines += $currentLine
            }
            $button.Text = $lines -join "`n"
        } else {
            $button.Text = $buttonText
        }

        $button.Tag = @{
            ScriptPath = $toolScript.FullName
            ScriptName = $toolScript.BaseName
        }

        # Update button color based on script name
        switch -Regex ($toolScript.BaseName) {
            'AD|Active|Lockout' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(70, 130, 180)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(100, 160, 210)
            }
            'Defender|Security|Toolkit' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(220, 20, 60)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(250, 50, 90)
            }
            'GPO|Policy|Report' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(46, 139, 87)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(76, 169, 117)
            }
            'Print' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(255, 140, 0)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(255, 170, 30)
            }
            'User|Data|Copy' {
                $button.BackColor = [System.Drawing.Color]::FromArgb(75, 0, 130)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(105, 30, 160)
            }
            default {
                $button.BackColor = [System.Drawing.Color]::FromArgb(105, 105, 105)
                $button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(135, 135, 135)
            }
        }

        $button.Visible = $true
        $button.Enabled = $true
    }

    # Hide any unused buttons
    for ($i = $toolScripts.Count; $i -lt $script:ToolButtons.Count; $i++) {
        $script:ToolButtons[$i].Visible = $false
    }

    Write-Verbose "Updated $($toolScripts.Count) toolbox buttons"
}

# Export only the functions called from main script
Export-ModuleMember -Function @(
    'Initialize-ToolboxPanel',
    'Update-ToolboxButtons'
)