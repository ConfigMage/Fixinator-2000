# RemoteConnection.psm1
# Simplified remote connection helpers used by the Remote Help Desk Tool

$script:CredentialContext = [pscustomobject]@{
    Credential = $null
    AuthenticationMethod = 'Unknown'
}
$script:ConnectedComputers = @{}

function Set-RemoteCredential {
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential,

        [Parameter()]
        [ValidateSet('UsernamePassword','SmartCard','CurrentUser','Unknown')]
        [string]$AuthenticationMethod = 'Unknown'
    )

    if ($AuthenticationMethod -ne 'CurrentUser' -and $null -eq $Credential) {
        throw "Credential cannot be null when AuthenticationMethod is $AuthenticationMethod."
    }

    if ($AuthenticationMethod -eq 'CurrentUser') {
        $Credential = [System.Management.Automation.PSCredential]::Empty
    }

    $script:CredentialContext = [pscustomobject]@{
        Credential = $Credential
        AuthenticationMethod = $AuthenticationMethod
    }
}

function Get-RemoteCredential {
    [CmdletBinding()]
    param()

    if ($null -eq $script:CredentialContext) {
        return $null
    }

    $credential = $script:CredentialContext.Credential
    if ($credential -eq [System.Management.Automation.PSCredential]::Empty) {
        return $null
    }

    return $credential
}

function Get-RemoteAuthenticationMethod {
    [CmdletBinding()]
    param()

    if ($null -eq $script:CredentialContext) {
        return 'Unknown'
    }

    return $script:CredentialContext.AuthenticationMethod
}

function Get-RemoteAuthenticationSummary {
    [CmdletBinding()]
    param()

    switch (Get-RemoteAuthenticationMethod) {
        'SmartCard'        { 'Auth: Smart Card' }
        'CurrentUser'      { 'Auth: Current User' }
        'UsernamePassword' { 'Auth: Username/Password' }
        default            { 'Auth: Unknown' }
    }
}

function Invoke-WmiObjectWithCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$Namespace,

        [Parameter()]
        [string]$Query,

        [Parameter()]
        [string]$Class,

        [Parameter()]
        [string]$Filter,

        [Parameter()]
        [string[]]$Property
    )

    try {
        $credential = Get-RemoteCredential
        $params = @{
            ComputerName = $ComputerName
            Namespace    = $Namespace
            ErrorAction  = 'Stop'
        }

        if ($credential) { $params['Credential'] = $credential }
        if ($Query)   { $params['Query'] = $Query }
        if ($Class)   { $params['Class'] = $Class }
        if ($Filter)  { $params['Filter'] = $Filter }
        if ($Property){ $params['Property'] = $Property }

        return Get-WmiObject @params
    }
    catch {
        Write-Error "WMI query failed on $ComputerName : $_"
        return $null
    }
}

function Invoke-WmiMethodWithCredentials {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$Class,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter()]
        [hashtable]$ArgumentList
    )

    try {
        $credential = Get-RemoteCredential
        $params = @{
            ComputerName = $ComputerName
            Class        = $Class
            Name         = $Name
            ErrorAction  = 'Stop'
        }

        if ($credential) { $params['Credential'] = $credential }
        if ($ArgumentList) { $params['ArgumentList'] = $ArgumentList }

        return Invoke-WmiMethod @params
    }
    catch {
        Write-Error "WMI method invocation failed on $ComputerName : $_"
        return $null
    }
}

function Test-RemoteConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    $result = [ordered]@{
        ComputerName     = $ComputerName
        PingSucceeded    = $false
        WinRMSucceeded   = $false
        IsOnline         = $false
        IPAddress        = $null
        Error            = $null
        LastChecked      = Get-Date
    }

    try {
        $ping = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction Stop
        $result.PingSucceeded = $ping

        if ($ping) {
            try {
                $ip = (Test-Connection -ComputerName $ComputerName -Count 1 -ErrorAction Stop).IPv4Address.IPAddressToString
                if ($ip) { $result.IPAddress = $ip }
            }
            catch { }

            try {
                $session = New-PSSession -ComputerName $ComputerName -Credential (Get-RemoteCredential) -ErrorAction Stop
                if ($session) {
                    $result.WinRMSucceeded = $true
                    Remove-PSSession -Session $session
                }
            }
            catch {
                $result.Error = $_.Exception.Message
            }
        }

        $result.IsOnline = $result.PingSucceeded -and $result.WinRMSucceeded
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    $object = [pscustomobject]$result
    $script:ConnectedComputers[$ComputerName] = $object
    return $object
}

function Invoke-RemoteCommand {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [Parameter()]
        [object]$ArgumentList = $null,

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )

    try {
        if ($null -eq $Credential) {
            $Credential = Get-RemoteCredential
        }

        $connectionTest = Test-RemoteConnection -ComputerName $ComputerName
        if (-not $connectionTest.PingSucceeded) {
            throw "Computer $ComputerName is not reachable"
        }

        $sessionParams = @{
            ComputerName = $ComputerName
            ErrorAction  = 'Stop'
        }
        if ($Credential) { $sessionParams['Credential'] = $Credential }

        $session = New-PSSession @sessionParams

        try {
            $invocationScript = {
                param($remoteScript, $remoteArguments)

                if ($null -eq $remoteScript) {
                    throw 'No script block supplied for remote execution.'
                }

                if ($remoteScript -isnot [System.Management.Automation.ScriptBlock]) {
                    $remoteScript = [scriptblock]::Create($remoteScript.ToString())
                }

                if ($remoteArguments -is [System.Collections.IDictionary]) {
                    return & $remoteScript @remoteArguments
                }
                elseif ($null -eq $remoteArguments) {
                    return & $remoteScript
                }
                elseif ($remoteArguments -is [System.Collections.IEnumerable] -and -not ($remoteArguments -is [string])) {
                    return & $remoteScript @remoteArguments
                }
                else {
                    return & $remoteScript $remoteArguments
                }
            }

            $invokeArgs = @($ScriptBlock, $ArgumentList)
            $result = Invoke-Command -Session $session -ScriptBlock $invocationScript -ArgumentList $invokeArgs

            return @{
                Success = $true
                Result  = $result
            }
        }
        finally {
            if ($session) {
                Remove-PSSession -Session $session -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        return @{
            Success = $false
            Result  = $null
            Error   = $_.Exception.Message
        }
    }
}

function Invoke-RemoteScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [Parameter()]
        [hashtable]$Parameters = @{},

        [Parameter()]
        [System.Management.Automation.PSCredential]$Credential
    )

    if (-not (Test-Path $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }

    $scriptContent = Get-Content -Path $ScriptPath -Raw
    $scriptBlock = [scriptblock]::Create($scriptContent)

    $paramCopy = @{}
    if ($Parameters) {
        foreach ($key in $Parameters.Keys) {
            $paramCopy[$key] = $Parameters[$key]
        }
    }

    if (-not $paramCopy.ContainsKey('ComputerName')) {
        $paramCopy['ComputerName'] = $ComputerName
    }

    if ($Credential) {
        $paramCopy['Credential'] = $Credential
    }

    return Invoke-RemoteCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -ArgumentList $paramCopy -Credential $Credential
}

function Get-ConnectedComputers {
    [CmdletBinding()]
    param()

    return $script:ConnectedComputers.Values
}

function Clear-ConnectionCache {
    [CmdletBinding()]
    param()

    $script:ConnectedComputers = @{}
}

Export-ModuleMember -Function @(
    'Set-RemoteCredential',
    'Get-RemoteCredential',
    'Get-RemoteAuthenticationMethod',
    'Get-RemoteAuthenticationSummary',
    'Invoke-WmiObjectWithCredentials',
    'Invoke-WmiMethodWithCredentials',
    'Test-RemoteConnection',
    'Invoke-RemoteCommand',
    'Invoke-RemoteScript',
    'Get-ConnectedComputers',
    'Clear-ConnectionCache'
)
