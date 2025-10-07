# Helper Functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:logFile -Value $logEntry
}

function Show-ErrorMessage {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, "Error", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Error)
}

function Show-InfoMessage {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, "Information", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information)
}# Defender Toolkit PowerShell Application
# Created for DOR Desktop Engineering
# Version 1.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global variables
$script:credential = $null
$script:currentComputer = ""
$script:isConnected = $false
$script:logPath = "C:\Temp\DefenderToolkit\Logs"
$script:isDarkMode = $false

# Color Schemes
$script:lightColors = @{
    Background = [System.Drawing.Color]::White
    PanelBackground = [System.Drawing.Color]::WhiteSmoke
    ButtonPrimary = [System.Drawing.Color]::FromArgb(0, 51, 102)
    ButtonSecondary = [System.Drawing.Color]::Gray
    TextPrimary = [System.Drawing.Color]::Black
    TextSecondary = [System.Drawing.Color]::DarkGray
    StatusGood = [System.Drawing.Color]::LightGreen
    StatusBad = [System.Drawing.Color]::LightCoral
    Border = [System.Drawing.Color]::Gray
}

$script:darkColors = @{
    Background = [System.Drawing.Color]::FromArgb(30, 30, 30)
    PanelBackground = [System.Drawing.Color]::FromArgb(45, 45, 45)
    ButtonPrimary = [System.Drawing.Color]::FromArgb(0, 122, 204)
    ButtonSecondary = [System.Drawing.Color]::FromArgb(70, 70, 70)
    TextPrimary = [System.Drawing.Color]::White
    TextSecondary = [System.Drawing.Color]::LightGray
    StatusGood = [System.Drawing.Color]::FromArgb(0, 128, 0)
    StatusBad = [System.Drawing.Color]::FromArgb(220, 20, 60)
    Border = [System.Drawing.Color]::FromArgb(60, 60, 60)
}

$script:currentColors = $script:lightColors

# Create log directory if it doesn't exist
if (!(Test-Path $script:logPath)) {
    New-Item -ItemType Directory -Path $script:logPath -Force | Out-Null
}

# Initialize log file
$script:logFile = Join-Path $script:logPath "DefenderToolkit_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# ASR Rules Mapping Dictionary
$script:ASRRules = @{
    'D4F940AB-401B-4EFC-AADC-AD5F3C50688A' = 'Block credential stealing from Windows LSASS'
    '9E6C4E1F-7D60-472F-BA1A-A39EF669E4B2' = 'Block process creations from PSExec and WMI commands'
    'B2B3F03D-6A65-4F7B-A9C7-1C7EF74A9BA4' = 'Block untrusted and unsigned processes from USB'
    '92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B' = 'Block Win32 API calls from Office macros'
    '5BEB7EFE-FD9A-4556-801D-275E5FFC04CC' = 'Block executable content from email and webmail'
    'D3E037E1-3EB8-44C8-A917-57927947596D' = 'Block JavaScript or VBScript from launching downloaded content'
    '3B576869-A4EC-4529-8536-B80A7769E899' = 'Block Office applications from creating executable content'
    '75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84' = 'Block Office applications from injecting code into other processes'
    '26190899-1602-49E8-8B27-EB1D0A1CE869' = 'Block Office communication apps from creating child processes'
    '7674BA52-37EB-4A4F-A9A1-F0F9A1619A2C' = 'Block Adobe Reader from creating child processes'
    'E6DB77E5-3DF2-4CF1-B95A-636979351E5B' = 'Block persistence through WMI event subscription'
    '01443614-CD74-433A-B99E-2ECDC07BFC25' = 'Block executable files from running unless they meet criteria'
    'BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550' = 'Block executable content from email client and webmail'
    '41E8C7AA-3687-4579-B1A4-3D4E7C7B6E22' = 'Block JavaScript and VBScript from launching downloaded executable content'
    'C1DB55AB-C21A-4637-BB3F-A12568109D35' = 'Use advanced protection against ransomware'
    'A8F5898E-1DC8-49A9-9878-85004B8A61E6' = 'Block Webshell creation for servers'
}

# Theme Toggle Function
function Toggle-DarkMode {
    $script:isDarkMode = -not $script:isDarkMode
    $script:currentColors = if ($script:isDarkMode) { $script:darkColors } else { $script:lightColors }
    
    # Update form colors
    $form.BackColor = $script:currentColors.Background
    $leftPanel.BackColor = $script:currentColors.PanelBackground
    $mainPanel.BackColor = $script:currentColors.PanelBackground
    
    # Update output box
    $outputBox.BackColor = $script:currentColors.Background
    $outputBox.ForeColor = $script:currentColors.TextPrimary
    
    # Update labels
    $titleLabel.ForeColor = $script:currentColors.ButtonPrimary
    $footerLabel.ForeColor = $script:currentColors.TextSecondary
    $computerLabel.ForeColor = $script:currentColors.TextPrimary
    
    # Update text box
    $computerTextBox.BackColor = $script:currentColors.Background
    $computerTextBox.ForeColor = $script:currentColors.TextPrimary
    
    # Update group boxes
    foreach ($groupBox in @($connectionGroupBox, $statusGroupBox, $quickActionsGroupBox, $utilitiesGroupBox, $operationsGroupBox)) {
        $groupBox.ForeColor = $script:currentColors.TextPrimary
    }
    
    # Update buttons
    foreach ($button in @($connectButton, $refreshAllButton, $exportReportButton, $viewLogButton, 
                         $checkAVButton, $listExclusionsButton, $listASRButton, 
                         $cfaStatusButton, $tamperProtectionButton, $viewEventsButton)) {
        if ($button.Enabled) {
            $button.BackColor = $script:currentColors.ButtonPrimary
        }
    }
    
    # Update status panel colors
    Update-StatusPanel -Connected $script:isConnected -ComputerName $script:currentComputer
    
    # Update dark mode button text
    $darkModeButton.Text = if ($script:isDarkMode) { "Light Mode ☀" } else { "Dark Mode 🌙" }
    $darkModeButton.BackColor = $script:currentColors.ButtonSecondary
}

# Updated Update-StatusPanel function to use theme colors
function Update-StatusPanel {
    param([bool]$Connected, [string]$ComputerName = "")
    
    if ($Connected) {
        $statusPanel.BackColor = $script:currentColors.StatusGood
        $statusLabel.Text = "Connected to:`n$ComputerName"
        $statusLabel.ForeColor = if ($script:isDarkMode) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::DarkGreen }
        $lastRefreshLabel.Text = "Last Refresh:`n$(Get-Date -Format 'HH:mm:ss')"
    } else {
        $statusPanel.BackColor = $script:currentColors.StatusBad
        $statusLabel.Text = "Disconnected"
        $statusLabel.ForeColor = if ($script:isDarkMode) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::DarkRed }
        $lastRefreshLabel.Text = ""
    }
    $lastRefreshLabel.ForeColor = $script:currentColors.TextSecondary
}

function Enable-OperationButtons {
    param([bool]$Enable)
    
    $checkAVButton.Enabled = $Enable
    $listExclusionsButton.Enabled = $Enable
    $listASRButton.Enabled = $Enable
    $cfaStatusButton.Enabled = $Enable
    $tamperProtectionButton.Enabled = $Enable
    $viewEventsButton.Enabled = $Enable
    $exportReportButton.Enabled = $Enable
    $refreshAllButton.Enabled = $Enable
}

function Append-Output {
    param(
        [string]$Text,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::Black
    )
    
    $outputBox.SelectionStart = $outputBox.TextLength
    $outputBox.SelectionLength = 0
    $outputBox.SelectionColor = $Color
    $outputBox.AppendText($Text + "`r`n")
    $outputBox.ScrollToCaret()
}

# Remote Command Execution Function
function Invoke-RemoteCommand {
    param(
        [string]$ComputerName,
        [scriptblock]$ScriptBlock,
        [System.Management.Automation.PSCredential]$Credential
    )
    
    try {
        Write-Log "Executing remote command on $ComputerName"
        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $ScriptBlock -Credential $Credential -ErrorAction Stop
        return $result
    }
    catch {
        Write-Log "Error executing remote command: $_" -Level "ERROR"
        throw $_
    }
}

# Connection Function
function Connect-RemoteComputer {
    $computerName = $computerTextBox.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        Show-ErrorMessage "Please enter a computer name"
        return
    }
    
    try {
        # Show connecting message
        Append-Output "Connecting to $computerName..." -Color ([System.Drawing.Color]::Blue)
        $outputBox.Refresh()
        
        # Get credentials
        $script:credential = Get-Credential -Message "Enter credentials for remote connection"
        
        if ($null -eq $script:credential) {
            Append-Output "Connection cancelled by user" -Color ([System.Drawing.Color]::Orange)
            return
        }
        
        # Test connection
        Write-Log "Testing connection to $computerName"
        $testConnection = Test-Connection -ComputerName $computerName -Count 1 -Quiet
        
        if (-not $testConnection) {
            throw "Unable to reach $computerName"
        }
        
        # Test PowerShell remoting
        $test = Invoke-RemoteCommand -ComputerName $computerName -ScriptBlock { $env:COMPUTERNAME } -Credential $script:credential
        
        $script:currentComputer = $computerName
        $script:isConnected = $true
        
        Update-StatusPanel -Connected $true -ComputerName $computerName
        Enable-OperationButtons -Enable $true
        
        Append-Output "Successfully connected to $computerName" -Color ([System.Drawing.Color]::Green)
        Write-Log "Successfully connected to $computerName"
        
        # Auto-check AV status
        Check-AVStatus
    }
    catch {
        $script:isConnected = $false
        Update-StatusPanel -Connected $false
        Enable-OperationButtons -Enable $false
        
        $errorMessage = "Failed to connect to $computerName`n$($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
        Show-ErrorMessage $errorMessage
    }
}

# AV Status Check Function
function Check-AVStatus {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Checking AV Status ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $avStatus = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            Get-MpComputerStatus | Select-Object AntivirusEnabled, RealTimeProtectionEnabled, 
                BehaviorMonitorEnabled, AMServiceEnabled, AntivirusSignatureLastUpdated,
                AMEngineVersion, AntivirusSignatureVersion, AntispywareSignatureVersion
        }
        
        if ($avStatus) {
            Append-Output "Antivirus Enabled: $($avStatus.AntivirusEnabled)" -Color $(if ($avStatus.AntivirusEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red })
            Append-Output "Real-Time Protection: $($avStatus.RealTimeProtectionEnabled)" -Color $(if ($avStatus.RealTimeProtectionEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red })
            Append-Output "Behavior Monitor: $($avStatus.BehaviorMonitorEnabled)" -Color $(if ($avStatus.BehaviorMonitorEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red })
            Append-Output "AM Service: $($avStatus.AMServiceEnabled)" -Color $(if ($avStatus.AMServiceEnabled) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red })
            Append-Output "`nUpdate Information:" -Color ([System.Drawing.Color]::DarkGreen)
            Append-Output "Last Signature Update: $($avStatus.AntivirusSignatureLastUpdated)"
            Append-Output "Engine Version: $($avStatus.AMEngineVersion)"
            Append-Output "Antivirus Signature Version: $($avStatus.AntivirusSignatureVersion)"
            Append-Output "Antispyware Signature Version: $($avStatus.AntispywareSignatureVersion)"
        }
        
        Write-Log "AV Status check completed successfully"
    }
    catch {
        $errorMessage = "Failed to check AV status: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# List Exclusions Function
function List-Exclusions {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Defender Exclusions ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $exclusions = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            Get-MpPreference | Select-Object ExclusionPath, ExclusionProcess, ExclusionExtension
        }
        
        if ($exclusions.ExclusionPath) {
            Append-Output "`nPath Exclusions:" -Color ([System.Drawing.Color]::DarkGreen)
            foreach ($path in $exclusions.ExclusionPath) {
                Append-Output "  • $path"
            }
        } else {
            Append-Output "`nPath Exclusions: None" -Color ([System.Drawing.Color]::Gray)
        }
        
        if ($exclusions.ExclusionProcess) {
            Append-Output "`nProcess Exclusions:" -Color ([System.Drawing.Color]::DarkGreen)
            foreach ($process in $exclusions.ExclusionProcess) {
                Append-Output "  • $process"
            }
        } else {
            Append-Output "`nProcess Exclusions: None" -Color ([System.Drawing.Color]::Gray)
        }
        
        if ($exclusions.ExclusionExtension) {
            Append-Output "`nExtension Exclusions:" -Color ([System.Drawing.Color]::DarkGreen)
            foreach ($ext in $exclusions.ExclusionExtension) {
                Append-Output "  • $ext"
            }
        } else {
            Append-Output "`nExtension Exclusions: None" -Color ([System.Drawing.Color]::Gray)
        }
        
        Write-Log "Exclusions list retrieved successfully"
    }
    catch {
        $errorMessage = "Failed to list exclusions: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# List ASR Rules Function
function List-ASRRules {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Attack Surface Reduction Rules ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $asrData = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            $prefs = Get-MpPreference
            [PSCustomObject]@{
                Ids = $prefs.AttackSurfaceReductionRules_Ids
                Actions = $prefs.AttackSurfaceReductionRules_Actions
            }
        }
        
        if ($asrData.Ids -and $asrData.Ids.Count -gt 0) {
            for ($i = 0; $i -lt $asrData.Ids.Count; $i++) {
                $ruleId = $asrData.Ids[$i]
                $action = switch ($asrData.Actions[$i]) {
                    0 { "Disabled" }
                    1 { "Block" }
                    2 { "Audit" }
                    6 { "Warn" }
                    default { "Unknown" }
                }
                
                $ruleName = if ($script:ASRRules.ContainsKey($ruleId)) { 
                    $script:ASRRules[$ruleId] 
                } else { 
                    "Unknown Rule" 
                }
                
                $color = switch ($action) {
                    "Block" { [System.Drawing.Color]::Green }
                    "Audit" { [System.Drawing.Color]::Orange }
                    "Warn" { [System.Drawing.Color]::Yellow }
                    "Disabled" { [System.Drawing.Color]::Gray }
                    default { [System.Drawing.Color]::Black }
                }
                
                Append-Output "`nRule: $ruleName" -Color ([System.Drawing.Color]::DarkGreen)
                Append-Output "GUID: $ruleId"
                Append-Output "Action: $action" -Color $color
            }
        } else {
            Append-Output "No ASR rules configured" -Color ([System.Drawing.Color]::Gray)
        }
        
        Write-Log "ASR rules retrieved successfully"
    }
    catch {
        $errorMessage = "Failed to list ASR rules: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# CFA Status Function
function Check-CFAStatus {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Controlled Folder Access Status ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $cfaData = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            $prefs = Get-MpPreference
            [PSCustomObject]@{
                CFAEnabled = $prefs.EnableControlledFolderAccess
                AllowedApps = $prefs.ControlledFolderAccessAllowedApplications
                ProtectedFolders = $prefs.ControlledFolderAccessProtectedFolders
            }
        }
        
        $cfaStatus = switch ($cfaData.CFAEnabled) {
            0 { "Disabled" }
            1 { "Enabled" }
            2 { "Audit Mode" }
            default { "Unknown" }
        }
        
        $color = switch ($cfaStatus) {
            "Enabled" { [System.Drawing.Color]::Green }
            "Audit Mode" { [System.Drawing.Color]::Orange }
            "Disabled" { [System.Drawing.Color]::Red }
            default { [System.Drawing.Color]::Black }
        }
        
        Append-Output "CFA Status: $cfaStatus" -Color $color
        
        if ($cfaData.AllowedApps) {
            Append-Output "`nAllowed Applications:" -Color ([System.Drawing.Color]::DarkGreen)
            foreach ($app in $cfaData.AllowedApps) {
                Append-Output "  • $app"
            }
        } else {
            Append-Output "`nAllowed Applications: None" -Color ([System.Drawing.Color]::Gray)
        }
        
        if ($cfaData.ProtectedFolders) {
            Append-Output "`nProtected Folders:" -Color ([System.Drawing.Color]::DarkGreen)
            foreach ($folder in $cfaData.ProtectedFolders) {
                Append-Output "  • $folder"
            }
        } else {
            Append-Output "`nProtected Folders: Default Windows folders" -Color ([System.Drawing.Color]::Gray)
        }
        
        Write-Log "CFA status retrieved successfully"
    }
    catch {
        $errorMessage = "Failed to check CFA status: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# Tamper Protection Function
function Check-TamperProtection {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Tamper Protection Status ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $tamperStatus = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            try {
                $regPath = "HKLM:\SOFTWARE\Microsoft\Windows Defender\Features"
                $tamperProtection = Get-ItemProperty -Path $regPath -Name "TamperProtection" -ErrorAction Stop
                return $tamperProtection.TamperProtection
            }
            catch {
                return "Unable to read registry"
            }
        }
        
        if ($tamperStatus -eq "Unable to read registry") {
            Append-Output "Tamper Protection: Unable to read (may require elevated permissions)" -Color ([System.Drawing.Color]::Orange)
        } else {
            $status = if ($tamperStatus -eq 1) { "Enabled" } else { "Disabled" }
            $color = if ($tamperStatus -eq 1) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Red }
            Append-Output "Tamper Protection: $status" -Color $color
        }
        
        Write-Log "Tamper protection check completed"
    }
    catch {
        $errorMessage = "Failed to check tamper protection: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# View Controlled Folder Access Events Function
function View-DefenderEvents {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        Append-Output "`n=== Retrieving Last 25 Controlled Folder Access Events ===" -Color ([System.Drawing.Color]::DarkBlue)
        
        $events = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            try {
                # Focus specifically on CFA events (Event IDs 1123 and 1124)
                $maxEvents = 25
                
                $filterXml = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-Windows Defender/Operational">
    <Select Path="Microsoft-Windows-Windows Defender/Operational">
      *[System[(EventID=1123 or EventID=1124)]]
    </Select>
  </Query>
</QueryList>
"@
                
                Get-WinEvent -FilterXml $filterXml -MaxEvents $maxEvents -ErrorAction Stop | 
                    ForEach-Object {
                        $event = $_
                        $processPath = ""
                        $folderPath = ""
                        $fileName = ""
                        $initiatingProcessName = ""
                        $initiatingProcessFolder = ""
                        
                        # Extract detailed CFA information from the message
                        if ($event.Message -match "Process Name:\s*([^\r\n]+)") {
                            $processPath = $Matches[1].Trim()
                            $initiatingProcessName = Split-Path -Path $processPath -Leaf
                            $initiatingProcessFolder = Split-Path -Path $processPath -Parent
                        }
                        
                        if ($event.Message -match "Path:\s*([^\r\n]+)") {
                            $fullPath = $Matches[1].Trim()
                            $fileName = Split-Path -Path $fullPath -Leaf
                            $folderPath = Split-Path -Path $fullPath -Parent
                        }
                        
                        [PSCustomObject]@{
                            Timestamp = $event.TimeCreated
                            DeviceName = $env:COMPUTERNAME
                            FileName = $fileName
                            FolderPath = $folderPath
                            InitiatingProcessFileName = $initiatingProcessName
                            InitiatingProcessFolderPath = $initiatingProcessFolder
                            ActionType = if ($event.Id -eq 1123) { "ControlledFolderAccessViolationBlocked" } else { "ControlledFolderAccessViolationAudited" }
                            EventId = $event.Id
                            FullMessage = $event.Message
                        }
                    } | Sort-Object Timestamp -Descending
            }
            catch {
                Write-Error "Failed to get CFA events: $_"
                return @()
            }
        }
        
        if ($events -and $events.Count -gt 0) {
            foreach ($event in $events) {
                $color = if ($event.ActionType -eq "ControlledFolderAccessViolationBlocked") { 
                    [System.Drawing.Color]::Red 
                } else { 
                    [System.Drawing.Color]::Orange 
                }
                
                Append-Output "`n=================================================================================" -Color ([System.Drawing.Color]::Gray)
                Append-Output "[$($event.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))] $($event.ActionType)" -Color $color
                Append-Output "=================================================================================" -Color ([System.Drawing.Color]::Gray)
                
                Append-Output "Device Name:                   $($event.DeviceName)" -Color ([System.Drawing.Color]::DarkBlue)
                Append-Output "File Name:                     $($event.FileName)" -Color ([System.Drawing.Color]::DarkRed)
                Append-Output "Folder Path:                   $($event.FolderPath)" -Color ([System.Drawing.Color]::DarkRed)
                Append-Output "Initiating Process:            $($event.InitiatingProcessFileName)" -Color ([System.Drawing.Color]::DarkBlue)
                Append-Output "Initiating Process Folder:    $($event.InitiatingProcessFolderPath)" -Color ([System.Drawing.Color]::DarkBlue)
                
                # Show a summary of the full message if available
                if ($event.FullMessage) {
                    $message = $event.FullMessage -replace "`r`n", " " -replace "  +", " "
                    if ($message.Length -gt 200) {
                        $message = $message.Substring(0, 200) + "..."
                    }
                    Append-Output "`nFull Event Details:" -Color ([System.Drawing.Color]::DarkGray)
                    
                    # Wrap long messages
                    $words = $message -split ' '
                    $line = ""
                    foreach ($word in $words) {
                        if (($line + " " + $word).Length -gt 80) {
                            if ($line) { Append-Output "  $line" -Color ([System.Drawing.Color]::Gray) }
                            $line = $word
                        } else {
                            $line = if ($line) { "$line $word" } else { $word }
                        }
                    }
                    if ($line) { Append-Output "  $line" -Color ([System.Drawing.Color]::Gray) }
                }
            }
            
            Append-Output "`nTotal events displayed: $($events.Count)" -Color ([System.Drawing.Color]::Blue)
        } else {
            Append-Output "No events found matching the selected criteria" -Color ([System.Drawing.Color]::Gray)
        }
        
        Write-Log "Retrieved $($events.Count) Defender events"
    }
    catch {
        $errorMessage = "Failed to retrieve events: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
    }
}

# Refresh All Function
function Refresh-AllChecks {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    Append-Output "`n=== Refreshing All Checks ===" -Color ([System.Drawing.Color]::DarkBlue)
    
    Check-AVStatus
    List-Exclusions
    List-ASRRules
    Check-CFAStatus
    Check-TamperProtection
    
    Update-StatusPanel -Connected $true -ComputerName $script:currentComputer
    Append-Output "`n=== All Checks Completed ===" -Color ([System.Drawing.Color]::DarkGreen)
}

# Export Report Function
function Export-HTMLReport {
    if (-not $script:isConnected) {
        Show-ErrorMessage "Not connected to any computer"
        return
    }
    
    try {
        $exportPath = "C:\Temp\DefenderToolkit\Reports"
        if (!(Test-Path $exportPath)) {
            New-Item -ItemType Directory -Path $exportPath -Force | Out-Null
        }
        
        $reportFile = Join-Path $exportPath "DefenderReport_$($script:currentComputer)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        
        # Gather all data
        $avStatus = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            Get-MpComputerStatus
        }
        
        $preferences = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
            Get-MpPreference
        }
        
        # Create HTML report
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Defender Status Report - $($script:currentComputer)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #003366; color: white; padding: 20px; border-radius: 5px; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .status-good { color: green; font-weight: bold; }
        .status-bad { color: red; font-weight: bold; }
        .status-warning { color: orange; font-weight: bold; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; }
        .footer { margin-top: 30px; padding: 10px; text-align: center; color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class="header">
        <h1>Defender Status Report</h1>
        <p>Computer: $($script:currentComputer) | Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    </div>
    
    <div class="section">
        <h2>Antivirus Status</h2>
        <table>
            <tr>
                <th>Component</th>
                <th>Status</th>
            </tr>
            <tr>
                <td>Antivirus Enabled</td>
                <td class="$(if ($avStatus.AntivirusEnabled) {'status-good'} else {'status-bad'})">$($avStatus.AntivirusEnabled)</td>
            </tr>
            <tr>
                <td>Real-Time Protection</td>
                <td class="$(if ($avStatus.RealTimeProtectionEnabled) {'status-good'} else {'status-bad'})">$($avStatus.RealTimeProtectionEnabled)</td>
            </tr>
            <tr>
                <td>Behavior Monitor</td>
                <td class="$(if ($avStatus.BehaviorMonitorEnabled) {'status-good'} else {'status-bad'})">$($avStatus.BehaviorMonitorEnabled)</td>
            </tr>
            <tr>
                <td>AM Service</td>
                <td class="$(if ($avStatus.AMServiceEnabled) {'status-good'} else {'status-bad'})">$($avStatus.AMServiceEnabled)</td>
            </tr>
            <tr>
                <td>Last Signature Update</td>
                <td>$($avStatus.AntivirusSignatureLastUpdated)</td>
            </tr>
        </table>
    </div>
    
    <div class="section">
        <h2>Exclusions</h2>
        <h3>Path Exclusions</h3>
        $(if ($preferences.ExclusionPath) {
            "<ul>" + ($preferences.ExclusionPath | ForEach-Object { "<li>$_</li>" }) -join "" + "</ul>"
        } else {
            "<p>None configured</p>"
        })
        
        <h3>Process Exclusions</h3>
        $(if ($preferences.ExclusionProcess) {
            "<ul>" + ($preferences.ExclusionProcess | ForEach-Object { "<li>$_</li>" }) -join "" + "</ul>"
        } else {
            "<p>None configured</p>"
        })
        
        <h3>Extension Exclusions</h3>
        $(if ($preferences.ExclusionExtension) {
            "<ul>" + ($preferences.ExclusionExtension | ForEach-Object { "<li>$_</li>" }) -join "" + "</ul>"
        } else {
            "<p>None configured</p>"
        })
    </div>
    
    <div class="section">
        <h2>Attack Surface Reduction Rules</h2>
        <table>
            <tr>
                <th>Rule</th>
                <th>GUID</th>
                <th>Action</th>
            </tr>
"@
        
        if ($preferences.AttackSurfaceReductionRules_Ids) {
            for ($i = 0; $i -lt $preferences.AttackSurfaceReductionRules_Ids.Count; $i++) {
                $ruleId = $preferences.AttackSurfaceReductionRules_Ids[$i]
                $action = switch ($preferences.AttackSurfaceReductionRules_Actions[$i]) {
                    0 { "Disabled" }
                    1 { "Block" }
                    2 { "Audit" }
                    6 { "Warn" }
                    default { "Unknown" }
                }
                
                $ruleName = if ($script:ASRRules.ContainsKey($ruleId)) { 
                    $script:ASRRules[$ruleId] 
                } else { 
                    "Unknown Rule" 
                }
                
                $actionClass = switch ($action) {
                    "Block" { "status-good" }
                    "Audit" { "status-warning" }
                    "Disabled" { "status-bad" }
                    default { "" }
                }
                
                $html += @"
            <tr>
                <td>$ruleName</td>
                <td>$ruleId</td>
                <td class="$actionClass">$action</td>
            </tr>
"@
            }
        } else {
            $html += @"
            <tr>
                <td colspan="3">No ASR rules configured</td>
            </tr>
"@
        }
        
        $html += @"
        </table>
    </div>
"@
        
        # Add CFA events section
        try {
            $events = Invoke-RemoteCommand -ComputerName $script:currentComputer -Credential $script:credential -ScriptBlock {
                try {
                    # Focus specifically on CFA events for HTML report
                    $maxEvents = 25
                    
                    $filterXml = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-Windows Defender/Operational">
    <Select Path="Microsoft-Windows-Windows Defender/Operational">
      *[System[(EventID=1123 or EventID=1124)]]
    </Select>
  </Query>
</QueryList>
"@
                    
                    Get-WinEvent -FilterXml $filterXml -MaxEvents $maxEvents -ErrorAction Stop | 
                        ForEach-Object {
                            $event = $_
                            $processPath = ""
                            $folderPath = ""
                            $fileName = ""
                            $initiatingProcessName = ""
                            $initiatingProcessFolder = ""
                            
                            # Extract detailed CFA information from the message
                            if ($event.Message -match "Process Name:\s*([^\r\n]+)") {
                                $processPath = $Matches[1].Trim()
                                $initiatingProcessName = Split-Path -Path $processPath -Leaf
                                $initiatingProcessFolder = Split-Path -Path $processPath -Parent
                            }
                            
                            if ($event.Message -match "Path:\s*([^\r\n]+)") {
                                $fullPath = $Matches[1].Trim()
                                $fileName = Split-Path -Path $fullPath -Leaf
                                $folderPath = Split-Path -Path $fullPath -Parent
                            }
                            
                            [PSCustomObject]@{
                                Timestamp = $event.TimeCreated
                                DeviceName = $env:COMPUTERNAME
                                FileName = $fileName
                                FolderPath = $folderPath
                                InitiatingProcessFileName = $initiatingProcessName
                                InitiatingProcessFolderPath = $initiatingProcessFolder
                                ActionType = if ($event.Id -eq 1123) { "ControlledFolderAccessViolationBlocked" } else { "ControlledFolderAccessViolationAudited" }
                                EventId = $event.Id
                            }
                        } | Sort-Object Timestamp -Descending
                }
                catch {
                    Write-Error "Failed to get CFA events for report: $_"
                    return @()
                }
            }
            
            if ($events -and $events.Count -gt 0) {
                Write-Log "Processing $($events.Count) CFA events for HTML report"
                $html += @"
    
    <div class="section">
        <h2>Controlled Folder Access Events (Last $($events.Count) events)</h2>
        <table>
            <tr>
                <th>Timestamp</th>
                <th>Action Type</th>
                <th>File Name</th>
                <th>Folder Path</th>
                <th>Initiating Process</th>
                <th>Process Folder</th>
            </tr>
"@
                
                $rowCount = 0
                foreach ($event in $events) {
                    $rowCount++
                    try {
                        # Provide fallback values for null/empty properties
                        $timestamp = if ($event.Timestamp) { $event.Timestamp.ToString('yyyy-MM-dd HH:mm:ss') } else { "Unknown" }
                        $actionType = if ($event.ActionType) { $event.ActionType } else { "Unknown" }
                        $fileName = if ($event.FileName) { $event.FileName } else { "N/A" }
                        $folderPath = if ($event.FolderPath) { $event.FolderPath } else { "N/A" }
                        $processName = if ($event.InitiatingProcessFileName) { $event.InitiatingProcessFileName } else { "N/A" }
                        $processFolder = if ($event.InitiatingProcessFolderPath) { $event.InitiatingProcessFolderPath } else { "N/A" }
                        
                        $rowClass = if ($actionType -eq "ControlledFolderAccessViolationBlocked") { 
                            "status-bad" 
                        } else { 
                            "status-warning" 
                        }
                        
                        $actionDisplay = if ($actionType -eq "ControlledFolderAccessViolationBlocked") {
                            "Blocked"
                        } else {
                            "Audited"
                        }
                        
                        # Safely encode HTML characters
                        $safeFileName = ($fileName -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
                        $safeFolderPath = ($folderPath -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
                        $safeProcessName = ($processName -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
                        $safeProcessFolder = ($processFolder -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
                        
                        Write-Log "Adding row $rowCount to HTML: $actionDisplay - $safeFileName"
                        
                        $html += @"
            <tr>
                <td style='white-space: nowrap;'>$timestamp</td>
                <td class="$rowClass">$actionDisplay</td>
                <td style='max-width: 150px; word-wrap: break-word;'>$safeFileName</td>
                <td style='max-width: 200px; word-wrap: break-word;'>$safeFolderPath</td>
                <td style='max-width: 150px; word-wrap: break-word;'>$safeProcessName</td>
                <td style='max-width: 200px; word-wrap: break-word;'>$safeProcessFolder</td>
            </tr>
"@
                    }
                    catch {
                        Write-Log "Error processing event $rowCount for HTML: $_" -Level "ERROR"
                        # Add a simple error row
                        $html += @"
            <tr>
                <td colspan="6" style="color: red;">Error processing event $rowCount</td>
            </tr>
"@
                    }
                }
                
                $html += @"
        </table>
    </div>
"@
            } else {
                $html += @"
    
    <div class="section">
        <h2>Controlled Folder Access Events</h2>
        <p>No Controlled Folder Access events found in the Windows Defender log.</p>
    </div>
"@
            }
        }
        catch {
            # If events fail, just continue without them
            Write-Log "Failed to include events in report: $_" -Level "WARNING"
        }
        
        $html += @"
    
    <div class="footer">
        <p>Generated by Defender Toolkit | Brought to you by DOR Desktop Engineering</p>
    </div>
</body>
</html>
"@
        
        $html | Out-File -FilePath $reportFile -Encoding UTF8
        
        Append-Output "`nReport exported to: $reportFile" -Color ([System.Drawing.Color]::Green)
        Show-InfoMessage "Report exported successfully to:`n$reportFile"
        Write-Log "Report exported to $reportFile"
    }
    catch {
        $errorMessage = "Failed to export report: $($_.Exception.Message)"
        Append-Output $errorMessage -Color ([System.Drawing.Color]::Red)
        Write-Log $errorMessage -Level "ERROR"
        Show-ErrorMessage $errorMessage
    }
}

# Create Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Defender Toolkit"
$form.Size = New-Object System.Drawing.Size(1200, 800)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.BackColor = [System.Drawing.Color]::White

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "DEFENDER TOOLKIT"
$titleLabel.Size = New-Object System.Drawing.Size(1180, 40)
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 51, 102)  # Dark blue
$titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($titleLabel)

# Left Panel
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Size = New-Object System.Drawing.Size(250, 680)
$leftPanel.Location = New-Object System.Drawing.Point(10, 60)
$leftPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($leftPanel)

# Connection Section
$connectionGroupBox = New-Object System.Windows.Forms.GroupBox
$connectionGroupBox.Text = "Connection"
$connectionGroupBox.Size = New-Object System.Drawing.Size(230, 90)
$connectionGroupBox.Location = New-Object System.Drawing.Point(10, 10)
$connectionGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($connectionGroupBox)

$computerLabel = New-Object System.Windows.Forms.Label
$computerLabel.Text = "Computer:"
$computerLabel.Size = New-Object System.Drawing.Size(70, 20)
$computerLabel.Location = New-Object System.Drawing.Point(10, 25)
$computerLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$connectionGroupBox.Controls.Add($computerLabel)

$computerTextBox = New-Object System.Windows.Forms.TextBox
$computerTextBox.Size = New-Object System.Drawing.Size(140, 22)
$computerTextBox.Location = New-Object System.Drawing.Point(80, 23)
$computerTextBox.Font = New-Object System.Drawing.Font("Arial", 9)
$connectionGroupBox.Controls.Add($computerTextBox)

$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Text = "Connect"
$connectButton.Size = New-Object System.Drawing.Size(210, 30)
$connectButton.Location = New-Object System.Drawing.Point(10, 52)
$connectButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$connectButton.ForeColor = [System.Drawing.Color]::White
$connectButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$connectButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$connectButton.Add_Click({ Connect-RemoteComputer })
$connectionGroupBox.Controls.Add($connectButton)

# Status Panel
$statusGroupBox = New-Object System.Windows.Forms.GroupBox
$statusGroupBox.Text = "Status"
$statusGroupBox.Size = New-Object System.Drawing.Size(230, 120)
$statusGroupBox.Location = New-Object System.Drawing.Point(10, 110)
$statusGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($statusGroupBox)

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Size = New-Object System.Drawing.Size(210, 60)
$statusPanel.Location = New-Object System.Drawing.Point(10, 20)
$statusPanel.BackColor = [System.Drawing.Color]::LightCoral
$statusPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$statusGroupBox.Controls.Add($statusPanel)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Disconnected"
$statusLabel.Size = New-Object System.Drawing.Size(200, 40)
$statusLabel.Location = New-Object System.Drawing.Point(5, 5)
$statusLabel.Font = New-Object System.Drawing.Font("Arial", 9)
$statusLabel.ForeColor = [System.Drawing.Color]::DarkRed
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusPanel.Controls.Add($statusLabel)

$lastRefreshLabel = New-Object System.Windows.Forms.Label
$lastRefreshLabel.Text = ""
$lastRefreshLabel.Size = New-Object System.Drawing.Size(210, 30)
$lastRefreshLabel.Location = New-Object System.Drawing.Point(10, 85)
$lastRefreshLabel.Font = New-Object System.Drawing.Font("Arial", 8)
$lastRefreshLabel.ForeColor = [System.Drawing.Color]::DarkGray
$lastRefreshLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$statusGroupBox.Controls.Add($lastRefreshLabel)

# Quick Actions
$quickActionsGroupBox = New-Object System.Windows.Forms.GroupBox
$quickActionsGroupBox.Text = "Quick Actions"
$quickActionsGroupBox.Size = New-Object System.Drawing.Size(230, 70)
$quickActionsGroupBox.Location = New-Object System.Drawing.Point(10, 240)
$quickActionsGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($quickActionsGroupBox)

$refreshAllButton = New-Object System.Windows.Forms.Button
$refreshAllButton.Text = "Refresh All"
$refreshAllButton.Size = New-Object System.Drawing.Size(210, 30)
$refreshAllButton.Location = New-Object System.Drawing.Point(10, 25)
$refreshAllButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$refreshAllButton.ForeColor = [System.Drawing.Color]::White
$refreshAllButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$refreshAllButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$refreshAllButton.Enabled = $false
$refreshAllButton.Add_Click({ Refresh-AllChecks })
$quickActionsGroupBox.Controls.Add($refreshAllButton)

# Export & Utilities
$utilitiesGroupBox = New-Object System.Windows.Forms.GroupBox
$utilitiesGroupBox.Text = "Export & Utilities"
$utilitiesGroupBox.Size = New-Object System.Drawing.Size(230, 150)
$utilitiesGroupBox.Location = New-Object System.Drawing.Point(10, 320)
$utilitiesGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($utilitiesGroupBox)

$exportReportButton = New-Object System.Windows.Forms.Button
$exportReportButton.Text = "Export HTML Report"
$exportReportButton.Size = New-Object System.Drawing.Size(210, 30)
$exportReportButton.Location = New-Object System.Drawing.Point(10, 25)
$exportReportButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$exportReportButton.ForeColor = [System.Drawing.Color]::White
$exportReportButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$exportReportButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$exportReportButton.Enabled = $false
$exportReportButton.Add_Click({ Export-HTMLReport })
$utilitiesGroupBox.Controls.Add($exportReportButton)

$viewLogButton = New-Object System.Windows.Forms.Button
$viewLogButton.Text = "View Log File"
$viewLogButton.Size = New-Object System.Drawing.Size(210, 30)
$viewLogButton.Location = New-Object System.Drawing.Point(10, 60)
$viewLogButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$viewLogButton.ForeColor = [System.Drawing.Color]::White
$viewLogButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$viewLogButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$viewLogButton.Add_Click({ 
    Start-Process notepad.exe -ArgumentList $script:logFile
})
$utilitiesGroupBox.Controls.Add($viewLogButton)

# Dark Mode Toggle Button
$darkModeButton = New-Object System.Windows.Forms.Button
$darkModeButton.Text = "Dark Mode 🌙"
$darkModeButton.Size = New-Object System.Drawing.Size(210, 30)
$darkModeButton.Location = New-Object System.Drawing.Point(10, 95)
$darkModeButton.BackColor = [System.Drawing.Color]::Gray
$darkModeButton.ForeColor = [System.Drawing.Color]::White
$darkModeButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$darkModeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$darkModeButton.Add_Click({ Toggle-DarkMode })
$utilitiesGroupBox.Controls.Add($darkModeButton)

# Main Operations Area
$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Size = New-Object System.Drawing.Size(920, 680)
$mainPanel.Location = New-Object System.Drawing.Point(270, 60)
$mainPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($mainPanel)

# Output Display
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Size = New-Object System.Drawing.Size(900, 400)
$outputBox.Location = New-Object System.Drawing.Point(10, 10)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
$outputBox.BackColor = [System.Drawing.Color]::White
$outputBox.ForeColor = [System.Drawing.Color]::Black
$mainPanel.Controls.Add($outputBox)

# Operation Buttons
$operationsGroupBox = New-Object System.Windows.Forms.GroupBox
$operationsGroupBox.Text = "Operations"
$operationsGroupBox.Size = New-Object System.Drawing.Size(900, 200)
$operationsGroupBox.Location = New-Object System.Drawing.Point(10, 420)
$operationsGroupBox.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$mainPanel.Controls.Add($operationsGroupBox)

# First row of buttons
$checkAVButton = New-Object System.Windows.Forms.Button
$checkAVButton.Text = "Check AV Status"
$checkAVButton.Size = New-Object System.Drawing.Size(280, 40)
$checkAVButton.Location = New-Object System.Drawing.Point(20, 30)
$checkAVButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$checkAVButton.ForeColor = [System.Drawing.Color]::White
$checkAVButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$checkAVButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$checkAVButton.Enabled = $false
$checkAVButton.Add_Click({ Check-AVStatus })
$operationsGroupBox.Controls.Add($checkAVButton)

$listExclusionsButton = New-Object System.Windows.Forms.Button
$listExclusionsButton.Text = "List Exclusions"
$listExclusionsButton.Size = New-Object System.Drawing.Size(280, 40)
$listExclusionsButton.Location = New-Object System.Drawing.Point(310, 30)
$listExclusionsButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$listExclusionsButton.ForeColor = [System.Drawing.Color]::White
$listExclusionsButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$listExclusionsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$listExclusionsButton.Enabled = $false
$listExclusionsButton.Add_Click({ List-Exclusions })
$operationsGroupBox.Controls.Add($listExclusionsButton)

$listASRButton = New-Object System.Windows.Forms.Button
$listASRButton.Text = "List ASR Rules"
$listASRButton.Size = New-Object System.Drawing.Size(280, 40)
$listASRButton.Location = New-Object System.Drawing.Point(600, 30)
$listASRButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$listASRButton.ForeColor = [System.Drawing.Color]::White
$listASRButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$listASRButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$listASRButton.Enabled = $false
$listASRButton.Add_Click({ List-ASRRules })
$operationsGroupBox.Controls.Add($listASRButton)

# Second row of buttons
$cfaStatusButton = New-Object System.Windows.Forms.Button
$cfaStatusButton.Text = "CFA Status"
$cfaStatusButton.Size = New-Object System.Drawing.Size(280, 40)
$cfaStatusButton.Location = New-Object System.Drawing.Point(20, 80)
$cfaStatusButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$cfaStatusButton.ForeColor = [System.Drawing.Color]::White
$cfaStatusButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$cfaStatusButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$cfaStatusButton.Enabled = $false
$cfaStatusButton.Add_Click({ Check-CFAStatus })
$operationsGroupBox.Controls.Add($cfaStatusButton)

$tamperProtectionButton = New-Object System.Windows.Forms.Button
$tamperProtectionButton.Text = "Tamper Protection"
$tamperProtectionButton.Size = New-Object System.Drawing.Size(280, 40)
$tamperProtectionButton.Location = New-Object System.Drawing.Point(310, 80)
$tamperProtectionButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$tamperProtectionButton.ForeColor = [System.Drawing.Color]::White
$tamperProtectionButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$tamperProtectionButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$tamperProtectionButton.Enabled = $false
$tamperProtectionButton.Add_Click({ Check-TamperProtection })
$operationsGroupBox.Controls.Add($tamperProtectionButton)

$viewEventsButton = New-Object System.Windows.Forms.Button
$viewEventsButton.Text = "View CFA Events"
$viewEventsButton.Size = New-Object System.Drawing.Size(280, 40)
$viewEventsButton.Location = New-Object System.Drawing.Point(600, 80)
$viewEventsButton.BackColor = [System.Drawing.Color]::FromArgb(0, 51, 102)
$viewEventsButton.ForeColor = [System.Drawing.Color]::White
$viewEventsButton.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$viewEventsButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$viewEventsButton.Enabled = $false
$viewEventsButton.Add_Click({ View-DefenderEvents })
$operationsGroupBox.Controls.Add($viewEventsButton)

# Admin functions row
$clearOutputButton = New-Object System.Windows.Forms.Button
$clearOutputButton.Text = "Clear Output"
$clearOutputButton.Size = New-Object System.Drawing.Size(280, 40)
$clearOutputButton.Location = New-Object System.Drawing.Point(310, 130)
$clearOutputButton.BackColor = [System.Drawing.Color]::Gray
$clearOutputButton.ForeColor = [System.Drawing.Color]::White
$clearOutputButton.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
$clearOutputButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$clearOutputButton.Add_Click({ 
    $outputBox.Clear()
    Append-Output "Output cleared" -Color ([System.Drawing.Color]::Gray)
})
$operationsGroupBox.Controls.Add($clearOutputButton)

# Footer
$footerLabel = New-Object System.Windows.Forms.Label
$footerLabel.Text = "Brought to you by DOR Desktop Engineering"
$footerLabel.Size = New-Object System.Drawing.Size(1180, 20)
$footerLabel.Location = New-Object System.Drawing.Point(10, 750)
$footerLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
$footerLabel.ForeColor = [System.Drawing.Color]::Gray
$footerLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($footerLabel)

# Initialize
Write-Log "Defender Toolkit started"
Append-Output "Defender Toolkit initialized. Please connect to a remote computer to begin." -Color ([System.Drawing.Color]::Blue)

# Show form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()

# Cleanup
Write-Log "Defender Toolkit closed"