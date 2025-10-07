#Requires -Version 5.1
# NOTE: Administrator privileges recommended for remote operations

<#
.SYNOPSIS
    The Fixinator 2000
.DESCRIPTION
    A comprehensive help desk toolkit with remote software installation, toolbox utilities, and script library
.AUTHOR
    Help Desk Team
.VERSION
    2000.0
#>

[CmdletBinding()]
param()

# Set strict mode for better error handling
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Get script directory
if ($PSScriptRoot) {
    $script:ScriptRoot = $PSScriptRoot
}
elseif ($PSCommandPath) {
    $script:ScriptRoot = Split-Path -Parent $PSCommandPath
}
elseif ($MyInvocation.ScriptName) {
    $script:ScriptRoot = Split-Path -Parent $MyInvocation.ScriptName
}
elseif ($MyInvocation.MyCommand -and $MyInvocation.MyCommand.PSObject.Properties['Path']) {
    $script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
else {
    $script:ScriptRoot = (Get-Location).ProviderPath
    Write-Warning "Unable to determine script root from invocation metadata; using current location: $script:ScriptRoot"
}

# Import required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define professional color scheme
$script:ColorScheme = @{
    # Primary colors
    PrimaryDark = [System.Drawing.Color]::FromArgb(25, 42, 86)      # Dark blue
    PrimaryMedium = [System.Drawing.Color]::FromArgb(41, 128, 185)  # Professional blue
    PrimaryLight = [System.Drawing.Color]::FromArgb(52, 152, 219)   # Light blue

    # Accent colors
    AccentGreen = [System.Drawing.Color]::FromArgb(39, 174, 96)     # Success green
    AccentOrange = [System.Drawing.Color]::FromArgb(230, 126, 34)   # Warning orange
    AccentRed = [System.Drawing.Color]::FromArgb(231, 76, 60)       # Error red

    # Neutral colors
    BackgroundLight = [System.Drawing.Color]::FromArgb(236, 240, 241) # Light gray background
    BackgroundWhite = [System.Drawing.Color]::White
    TextDark = [System.Drawing.Color]::FromArgb(44, 62, 80)         # Dark text
    TextLight = [System.Drawing.Color]::FromArgb(149, 165, 166)     # Light text
    BorderColor = [System.Drawing.Color]::FromArgb(189, 195, 199)   # Border gray

    # Menu and tab colors
    MenuBackground = [System.Drawing.Color]::FromArgb(248, 249, 250)
    TabSelected = [System.Drawing.Color]::White
    TabUnselected = [System.Drawing.Color]::FromArgb(236, 240, 241)
}


# Ensure previous module versions are unloaded so we always import the latest functions
foreach ($moduleName in @("ConfigManager","Logging","RemoteConnection","Toolbox","ScriptLibrary","SoftwareCenter")) {
    Remove-Module -Name $moduleName -ErrorAction SilentlyContinue
}

# Import modules
$moduleFiles = @(
    "ConfigManager.psm1",
    "Logging.psm1",
    "RemoteConnection.psm1",
    "Toolbox.psm1",
    "ScriptLibrary.psm1",
    "SoftwareCenter.psm1"
)

foreach ($module in $moduleFiles) {
    $modulePath = Join-Path -Path "$ScriptRoot\Modules" -ChildPath $module
    if (Test-Path $modulePath) {
        try {
            Import-Module $modulePath -Force -ErrorAction Stop
            Write-Verbose "Successfully imported module: $module"
        }
        catch {
            Write-Error "Failed to import module $module : $_"
            exit 1
        }
    }
    else {
        Write-Warning "Module not found: $modulePath"
    }
}

# Initialize logging
Initialize-Logging -LogPath "$ScriptRoot\Logs"
Write-Log "Starting The Fixinator 2000" -Level INFO

# Global variables to store form data
$script:FormData = @{
    ScriptRoot = $ScriptRoot
    Credential = $null
    AuthenticationMethod = 'Unknown'
    AuthenticationSummary = 'Auth: Unknown'
    AuthenticationWarningShown = $false
    StatusLabel = $null
    ConnectionLabel = $null
    SoftwareCenterTab = $null
    ToolboxTab = $null
    ScriptLibraryTab = $null
}

# Function to prompt for credentials
function Get-AuthenticationSummaryText {
    [CmdletBinding()]
    param()

    switch ($script:FormData.AuthenticationMethod) {
        'SmartCard' { 'Auth: Smart Card' }
        'CurrentUser' { 'Auth: Current User' }
        'UsernamePassword' { 'Auth: Username/Password' }
        default { 'Auth: Unknown' }
    }
}

function Get-UserCredential {
    [CmdletBinding()]
    param()

    $message = 'Authenticate to use The Fixinator 2000. Smart card users can insert their card and select it from the drop-down.'

    try {
        $credential = Get-Credential -Message $message -ErrorAction Stop
    }
    catch {
        Write-Log "Credential prompt failed: $_" -Level ERROR
        return $null
    }

    if ($null -eq $credential) {
        $script:FormData.AuthenticationMethod = 'Unknown'
        $script:FormData.AuthenticationSummary = Get-AuthenticationSummaryText
        return $null
    }

    if ($credential -eq [System.Management.Automation.PSCredential]::Empty) {
        $script:FormData.AuthenticationMethod = 'CurrentUser'
    }
    else {
        $password = $credential.GetNetworkCredential().Password
        if ([string]::IsNullOrWhiteSpace($password)) {
            $script:FormData.AuthenticationMethod = 'SmartCard'
        }
        else {
            $script:FormData.AuthenticationMethod = 'UsernamePassword'
        }
    }

    $script:FormData.AuthenticationSummary = Get-AuthenticationSummaryText
    return $credential
}

# Function to load configurations
function Invoke-LoadConfigurations {
    try {
        # Load Toolbox configuration (dynamic from ToolBox directory)
        Update-ToolboxButtons
        Write-Log "Loaded Toolbox configuration" -Level SUCCESS

        # Load Script Library configuration
        $scriptLibraryConfig = Get-ScriptLibraryConfig -ConfigPath (Join-Path $script:FormData.ScriptRoot "Config\ScriptLibrary.json")
        if ($scriptLibraryConfig) {
            Update-ScriptLibrary -Scripts $scriptLibraryConfig
            Write-Log "Loaded Script Library configuration" -Level SUCCESS
        }

        $script:FormData.StatusLabel.Text = "Configuration loaded successfully"
    }
    catch {
        Write-Log "Failed to load configuration: $_" -Level ERROR
        $script:FormData.StatusLabel.Text = "Configuration load failed"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load configuration files: $_",
            "Configuration Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    }
}

# Function to reload configuration
function Invoke-ReloadConfiguration {
    $script:FormData.StatusLabel.Text = "Reloading configuration..."
    Write-Log "Reloading configuration files" -Level INFO
    
    Invoke-LoadConfigurations
    
    [System.Windows.Forms.MessageBox]::Show(
        "Configuration files have been reloaded successfully.",
        "Configuration Reload",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Create the main form
function New-MainForm {
    # Create main form
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "The Fixinator 2000"
    $mainForm.Size = New-Object System.Drawing.Size(1200, 800)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.Icon = [System.Drawing.SystemIcons]::Information
    $mainForm.MinimumSize = New-Object System.Drawing.Size(900, 700)
    $mainForm.BackColor = $script:ColorScheme.BackgroundLight
    $mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    # Create menu bar
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.BackColor = $script:ColorScheme.MenuBackground
    $menuStrip.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $menuStrip.Padding = New-Object System.Windows.Forms.Padding(5, 2, 0, 2)

    # File menu
    $fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $fileMenu.Text = "&File"
    $fileMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    
    $reloadConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $reloadConfigItem.Text = "&Reload Configuration"
    $reloadConfigItem.ShortcutKeys = [System.Windows.Forms.Keys]::F5
    $reloadConfigItem.Add_Click({
        Invoke-ReloadConfiguration
    })
    
    $openConfigItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $openConfigItem.Text = "&Open Config Folder"
    $openConfigItem.Add_Click({
        Start-Process explorer.exe (Join-Path $script:FormData.ScriptRoot "Config")
    })
    
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "E&xit"
    $exitItem.ShortcutKeys = [System.Windows.Forms.Keys]::Alt, [System.Windows.Forms.Keys]::F4
    $exitItem.Add_Click({
        $mainForm.Close()
    })
    
    $separator = New-Object System.Windows.Forms.ToolStripSeparator
    $fileMenu.DropDownItems.AddRange(@($reloadConfigItem, $openConfigItem, $separator, $exitItem)) | Out-Null
    
    # Help menu
    $helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpMenu.Text = "&Help"
    $helpMenu.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutItem.Text = "&About"
    $aboutItem.Add_Click({
        # Create custom About dialog
        $aboutForm = New-Object System.Windows.Forms.Form
        $aboutForm.Text = "About The Fixinator 2000"
        $aboutForm.Size = New-Object System.Drawing.Size(650, 500)
        $aboutForm.StartPosition = "CenterParent"
        $aboutForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $aboutForm.MaximizeBox = $false
        $aboutForm.MinimizeBox = $false
        $aboutForm.BackColor = [System.Drawing.Color]::White
        $aboutForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

        # Title
        $lblTitle = New-Object System.Windows.Forms.Label
        $lblTitle.Text = "The Fixinator 2000"
        $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
        $lblTitle.ForeColor = $script:ColorScheme.PrimaryDark
        $lblTitle.Location = New-Object System.Drawing.Point(30, 20)
        $lblTitle.Size = New-Object System.Drawing.Size(580, 40)
        $aboutForm.Controls.Add($lblTitle)

        # Version
        $lblVersion = New-Object System.Windows.Forms.Label
        $lblVersion.Text = "Version 2000.0"
        $lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 10)
        $lblVersion.ForeColor = $script:ColorScheme.TextLight
        $lblVersion.Location = New-Object System.Drawing.Point(30, 60)
        $lblVersion.Size = New-Object System.Drawing.Size(580, 25)
        $aboutForm.Controls.Add($lblVersion)

        # Description (main text)
        $txtDescription = New-Object System.Windows.Forms.TextBox
        $txtDescription.Multiline = $true
        $txtDescription.ReadOnly = $true
        $txtDescription.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $txtDescription.BackColor = [System.Drawing.Color]::White
        $txtDescription.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $txtDescription.ForeColor = $script:ColorScheme.TextDark
        $txtDescription.Location = New-Object System.Drawing.Point(30, 95)
        $txtDescription.Size = New-Object System.Drawing.Size(580, 240)
        $txtDescription.Text = "The Fixinator 2000 is a comprehensive remote support toolkit designed to empower IT professionals with streamlined automation and deployment capabilities.`r`n`r`nBuilt during the 2025 DOR Hackathon by team `"The Last Picks`" (Andrew Fredrickson, Edgar Pozos, William Gorham, and Chris Solario), this tool combines three powerful features into one unified interface:`r`n`r`nSoftware Center - Deploy and install software remotely with ease`r`n`r`nPowerShell Toolbox - Access essential local utilities for system management`r`n`r`nScript Library - Create, store, and execute custom automation scripts"
        $aboutForm.Controls.Add($txtDescription)

        # Copyright
        $lblCopyright = New-Object System.Windows.Forms.Label
        $lblCopyright.Text = "© 2025 DOR Hackathon - Team `"The Last Picks`""
        $lblCopyright.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $lblCopyright.ForeColor = $script:ColorScheme.TextLight
        $lblCopyright.Location = New-Object System.Drawing.Point(30, 350)
        $lblCopyright.Size = New-Object System.Drawing.Size(580, 25)
        $lblCopyright.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $aboutForm.Controls.Add($lblCopyright)

        # Tagline
        $lblTagline = New-Object System.Windows.Forms.Label
        $lblTagline.Text = "Fixing problems since Y2K!"
        $lblTagline.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
        $lblTagline.ForeColor = $script:ColorScheme.PrimaryMedium
        $lblTagline.Location = New-Object System.Drawing.Point(30, 380)
        $lblTagline.Size = New-Object System.Drawing.Size(580, 25)
        $lblTagline.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $aboutForm.Controls.Add($lblTagline)

        # OK Button
        $btnOK = New-Object System.Windows.Forms.Button
        $btnOK.Text = "OK"
        $btnOK.Location = New-Object System.Drawing.Point(270, 415)
        $btnOK.Size = New-Object System.Drawing.Size(100, 32)
        $btnOK.BackColor = $script:ColorScheme.PrimaryMedium
        $btnOK.ForeColor = [System.Drawing.Color]::White
        $btnOK.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnOK.FlatAppearance.BorderSize = 0
        $btnOK.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $btnOK.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $aboutForm.Controls.Add($btnOK)
        $aboutForm.AcceptButton = $btnOK

        $aboutForm.ShowDialog() | Out-Null
    })

    $helpMenu.DropDownItems.Add($aboutItem) | Out-Null

    $menuStrip.Items.AddRange(@($fileMenu, $helpMenu)) | Out-Null
    $mainForm.Controls.Add($menuStrip)
    
    # Force layout calculation to get actual menu height
    $mainForm.PerformLayout()
    
    # Create TabControl with proper positioning
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControlY = $menuStrip.Height + 8
    $tabControlHeight = $mainForm.ClientSize.Height - $menuStrip.Height - 38  # Account for status bar
    $tabControl.Location = New-Object System.Drawing.Point(8, $tabControlY)
    $tabControl.Size = New-Object System.Drawing.Size(($mainForm.ClientSize.Width - 16), $tabControlHeight)
    $tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor
                          [System.Windows.Forms.AnchorStyles]::Bottom -bor
                          [System.Windows.Forms.AnchorStyles]::Left -bor
                          [System.Windows.Forms.AnchorStyles]::Right
    $tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $tabControl.Padding = New-Object System.Drawing.Point(15, 8)
    $tabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Normal
    
    # Create Software Center Tab
    $softwareCenterTab = New-Object System.Windows.Forms.TabPage
    $softwareCenterTab.Text = "  Software Center"
    $softwareCenterTab.BackColor = $script:ColorScheme.BackgroundWhite
    $softwareCenterTab.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:FormData.SoftwareCenterTab = $softwareCenterTab

    # Create Toolbox Tab
    $toolboxTab = New-Object System.Windows.Forms.TabPage
    $toolboxTab.Text = "  Toolbox"
    $toolboxTab.BackColor = $script:ColorScheme.BackgroundWhite
    $toolboxTab.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:FormData.ToolboxTab = $toolboxTab

    # Create Script Library Tab
    $scriptLibraryTab = New-Object System.Windows.Forms.TabPage
    $scriptLibraryTab.Text = "  Script Library"
    $scriptLibraryTab.BackColor = $script:ColorScheme.BackgroundWhite
    $scriptLibraryTab.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:FormData.ScriptLibraryTab = $scriptLibraryTab

    # Add tabs to control
    $tabControl.TabPages.AddRange(@($softwareCenterTab, $toolboxTab, $scriptLibraryTab)) | Out-Null
    $mainForm.Controls.Add($tabControl)
    
    # Create Status Strip
    $statusStrip = New-Object System.Windows.Forms.StatusStrip
    $statusStrip.BackColor = $script:ColorScheme.PrimaryDark
    $statusStrip.ForeColor = [System.Drawing.Color]::White
    $statusStrip.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $statusStrip.Padding = New-Object System.Windows.Forms.Padding(5, 0, 10, 0)

    $connectionLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $connectionLabel.Text = "Connection: Not Connected"
    $connectionLabel.BorderSides = [System.Windows.Forms.ToolStripStatusLabelBorderSides]::Right
    $connectionLabel.BorderStyle = [System.Windows.Forms.Border3DStyle]::Etched
    $connectionLabel.Width = 230
    $connectionLabel.ForeColor = [System.Drawing.Color]::White
    $connectionLabel.Margin = New-Object System.Windows.Forms.Padding(0, 3, 5, 2)
    $script:FormData.ConnectionLabel = $connectionLabel

    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = "Ready"
    $statusLabel.Spring = $true
    $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $statusLabel.ForeColor = [System.Drawing.Color]::White
    $statusLabel.Margin = New-Object System.Windows.Forms.Padding(5, 3, 0, 2)
    $script:FormData.StatusLabel = $statusLabel

    $statusStrip.Items.AddRange(@($connectionLabel, $statusLabel)) | Out-Null
    $mainForm.Controls.Add($statusStrip)
    
    # Form Load event
    $mainForm.Add_Load({
        $credential = Get-UserCredential

        if ($null -eq $credential) {
            [System.Windows.Forms.MessageBox]::Show(
                "Credentials are required to use this tool.",
                "Authentication Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $mainForm.Close()
            return
        }

        $script:FormData.Credential = $credential
        $setRemoteCommand = Get-Command Set-RemoteCredential -ErrorAction SilentlyContinue
        if ($setRemoteCommand -and $setRemoteCommand.Parameters.ContainsKey("AuthenticationMethod")) {
            Set-RemoteCredential -Credential $credential -AuthenticationMethod $script:FormData.AuthenticationMethod
        }
        elseif ($setRemoteCommand) {
            if (-not $script:FormData.AuthenticationWarningShown) {
                Write-Warning "Set-RemoteCredential does not support AuthenticationMethod; caching credential without method metadata."
                $script:FormData.AuthenticationWarningShown = $true
            }
            Set-RemoteCredential -Credential $credential
        }
        else {
            if (-not $script:FormData.AuthenticationWarningShown) {
                Write-Warning "Set-RemoteCredential command not found. Remote operations will prompt for credentials each time."
                $script:FormData.AuthenticationWarningShown = $true
            }
        }

        $authSummary = Get-AuthenticationSummaryText
        $script:FormData.AuthenticationSummary = $authSummary
        if ($script:FormData.ConnectionLabel) {
            $script:FormData.ConnectionLabel.Text = "Connection: Not Connected ($authSummary)"
        }

        # Initialize tabs
        Initialize-SoftwareCenterTab -TabPage $script:FormData.SoftwareCenterTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel

        # Initialize Toolbox tab
        Initialize-ToolboxPanel -Panel $script:FormData.ToolboxTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel -ConnectionLabel $script:FormData.ConnectionLabel

        # Initialize Script Library tab (now standalone)
        Initialize-ScriptLibraryPanel -Panel $script:FormData.ScriptLibraryTab -ScriptRoot $script:FormData.ScriptRoot -StatusLabel $script:FormData.StatusLabel

        # Load configurations
        Invoke-LoadConfigurations

        $script:FormData.StatusLabel.Text = "Ready"
        Write-Log "Form loaded successfully" -Level SUCCESS
    })
    
    # Form Closing event
    $mainForm.Add_FormClosing({
        Write-Log "Closing The Fixinator 2000" -Level INFO
    })
    
    # Ensure we return the form object explicitly
    return $mainForm
}

# Main execution
try {
    Write-Host "Starting The Fixinator 2000..." -ForegroundColor Green
    
    # Ensure required directories exist
    $requiredDirs = @("Logs", "Config", "Scripts", "Modules", "Resources")
    foreach ($dir in $requiredDirs) {
        $dirPath = Join-Path $ScriptRoot $dir
        if (!(Test-Path $dirPath)) {
            New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
            Write-Verbose "Created directory: $dirPath"
        }
    }
    
    # Create and show the main form
    $mainForm = New-MainForm
    [System.Windows.Forms.Application]::Run($mainForm)
}
catch {
    Write-Error "Fatal error: $_"
    Write-Log "Fatal error: $_" -Level ERROR
    [System.Windows.Forms.MessageBox]::Show(
        "A fatal error occurred:`n`n$_",
        "Fatal Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit 1
}
finally {
    # Cleanup
    Write-Log "Application closed" -Level INFO
}

