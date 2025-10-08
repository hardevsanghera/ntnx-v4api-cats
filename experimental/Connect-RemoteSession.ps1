#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Create a PowerShell session to a remote Windows host
    
.DESCRIPTION
    This script establishes a PowerShell remoting session to a remote Windows computer.
    It includes connection testing, credential management, and session management.
    
.PARAMETER ComputerName
    The name or IP address of the remote computer
    Default: 10.38.20.187
    
.PARAMETER Port
    The port to use for WinRM connection
    Default: 5985 (HTTP), 5986 (HTTPS)
    
.PARAMETER UseSSL
    Use HTTPS/SSL for the connection (port 5986)
    
.PARAMETER Credential
    PSCredential object for authentication. If not provided, will prompt for credentials
    
.PARAMETER TestConnection
    Test the connection before attempting to create session
    
.PARAMETER Interactive
    Enter an interactive session after connection
    
.EXAMPLE
    .\Connect-RemoteSession.ps1
    
.EXAMPLE
    .\Connect-RemoteSession.ps1 -ComputerName "server01" -TestConnection -Interactive
    
.EXAMPLE
    .\Connect-RemoteSession.ps1 -ComputerName "10.38.20.187" -UseSSL -Interactive
    
.NOTES
    - Requires PowerShell Remoting to be enabled on target computer
    - May require administrator privileges depending on configuration
    - Ensure firewall allows WinRM traffic (ports 5985/5986)
#>

[CmdletBinding()]
param(
    [string]$ComputerName = "10.38.20.187",
    [int]$Port = 0,  # Auto-select based on UseSSL
    [switch]$UseSSL,
    [PSCredential]$Credential,
    [switch]$TestConnection,
    [switch]$Interactive
)

# Script configuration
$ErrorActionPreference = 'Stop'

# Auto-select port if not specified
if ($Port -eq 0) {
    $Port = if ($UseSSL) { 5986 } else { 5985 }
}

Write-Host "🔗 PowerShell Remote Session Manager" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan
Write-Host ""

#region Helper Functions

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host "📦 $Title" -ForegroundColor Yellow
    Write-Host ("-" * ($Title.Length + 3)) -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "ℹ️  $Message" -ForegroundColor Blue
}

function Write-Warning {
    param([string]$Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

function Test-RemoteConnection {
    param(
        [string]$ComputerName,
        [int]$Port,
        [PSCredential]$Credential,
        [bool]$UseSSL
    )
    
    $results = @{}
    
    Write-Info "Testing connection to $ComputerName..."
    
    # Test basic network connectivity
    try {
        $pingResult = Test-NetConnection -ComputerName $ComputerName -Port $Port -WarningAction SilentlyContinue
        $results['NetworkConnectivity'] = if ($pingResult.TcpTestSucceeded) { "✅ Port $Port is open" } else { "❌ Port $Port is closed" }
    }
    catch {
        $results['NetworkConnectivity'] = "❌ Network test failed: $($_.Exception.Message)"
    }
    
    # Test WSMan
    try {
        $wsmanParams = @{
            ComputerName = $ComputerName
            ErrorAction = 'Stop'
        }
        
        if ($Credential) {
            $wsmanParams['Credential'] = $Credential
        }
        
        if ($UseSSL) {
            $wsmanParams['UseSSL'] = $true
        }
        
        Test-WSMan @wsmanParams | Out-Null
        $results['WSMan'] = "✅ WSMan service is available"
    }
    catch {
        $results['WSMan'] = "❌ WSMan test failed: $($_.Exception.Message)"
    }
    
    return $results
}

#endregion

#region Connection Information

Write-Section "Connection Information"

Write-Info "Target Computer: $ComputerName"
Write-Info "Port: $Port"
Write-Info "Protocol: $(if ($UseSSL) { 'HTTPS/SSL' } else { 'HTTP' })"
Write-Info "Authentication: $(if ($Credential) { $Credential.UserName } else { 'Will prompt for credentials' })"

#endregion

#region Credential Management

Write-Section "Credential Management"

if (-not $Credential) {
    Write-Info "No credentials provided. Prompting for authentication..."
    try {
        $Credential = Get-Credential -Message "Enter credentials for $ComputerName"
        Write-Success "Credentials obtained for user: $($Credential.UserName)"
    }
    catch {
        Write-Error "Failed to obtain credentials: $($_.Exception.Message)"
    }
}
else {
    Write-Success "Using provided credentials for user: $($Credential.UserName)"
}

#endregion

#region Connection Testing

if ($TestConnection) {
    Write-Section "Connection Testing"
    
    $testResults = Test-RemoteConnection -ComputerName $ComputerName -Port $Port -Credential $Credential -UseSSL:$UseSSL
    
    foreach ($test in $testResults.GetEnumerator()) {
        Write-Host "$($test.Key): $($test.Value)"
    }
    
    # Check if any tests failed
    $failedTests = $testResults.Values | Where-Object { $_ -like "❌*" }
    if ($failedTests) {
        Write-Warning "Some connection tests failed. Proceeding anyway..."
        Write-Info "If connection fails, check:"
        Write-Info "  • PowerShell Remoting is enabled: Enable-PSRemoting -Force"
        Write-Info "  • Firewall allows WinRM: Get-NetFirewallRule -Name '*WinRM*'"
        Write-Info "  • Trusted hosts configured if needed: Set-Item WSMan:\localhost\Client\TrustedHosts -Value '*'"
    }
    else {
        Write-Success "All connection tests passed!"
    }
}

#endregion

#region Session Creation

Write-Section "Creating PowerShell Session"

try {
    # Prepare session options
    $sessionParams = @{
        ComputerName = $ComputerName
        Credential = $Credential
        ErrorAction = 'Stop'
    }
    
    # Add SSL option if specified
    if ($UseSSL) {
        $sessionParams['UseSSL'] = $true
    }
    
    # Add port if not default
    if (($UseSSL -and $Port -ne 5986) -or (-not $UseSSL -and $Port -ne 5985)) {
        $sessionParams['Port'] = $Port
    }
    
    Write-Info "Attempting to create session with parameters:"
    $sessionParams.GetEnumerator() | ForEach-Object {
        if ($_.Key -ne 'Credential') {
            Write-Info "  $($_.Key): $($_.Value)"
        }
        else {
            Write-Info "  $($_.Key): [PSCredential for $($_.Value.UserName)]"
        }
    }
    
    # Create the session
    $session = New-PSSession @sessionParams
    
    if ($session) {
        Write-Success "PowerShell session created successfully!"
        Write-Info "Session ID: $($session.Id)"
        Write-Info "Session Name: $($session.Name)"
        Write-Info "Computer Name: $($session.ComputerName)"
        Write-Info "State: $($session.State)"
        
        # Store session in global variable for easy access
        $global:RemoteSession = $session
        Write-Info "Session stored in `$global:RemoteSession variable"
        
        # Test the session with a simple command
        Write-Info "Testing session with basic command..."
        try {
            $remoteInfo = Invoke-Command -Session $session -ScriptBlock {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
                    OSVersion = [System.Environment]::OSVersion.VersionString
                    LoggedOnUser = $env:USERNAME
                    CurrentTime = Get-Date
                }
            }
            
            Write-Success "Session test successful! Remote computer details:"
            Write-Info "  Computer Name: $($remoteInfo.ComputerName)"
            Write-Info "  PowerShell Version: $($remoteInfo.PowerShellVersion)"
            Write-Info "  OS Version: $($remoteInfo.OSVersion)"
            Write-Info "  Logged On User: $($remoteInfo.LoggedOnUser)"
            Write-Info "  Remote Time: $($remoteInfo.CurrentTime)"
        }
        catch {
            Write-Warning "Session created but test command failed: $($_.Exception.Message)"
        }
        
        # Interactive session option
        if ($Interactive) {
            Write-Section "Entering Interactive Session"
            Write-Info "Entering interactive session with $ComputerName..."
            Write-Info "Type 'exit' to return to local session"
            Write-Host ""
            
            try {
                Enter-PSSession -Session $session
            }
            catch {
                Write-Warning "Failed to enter interactive session: $($_.Exception.Message)"
            }
        }
        else {
            Write-Section "Session Management"
            Write-Info "Session is ready for use. You can:"
            Write-Info "  • Enter interactive mode: Enter-PSSession -Session `$global:RemoteSession"
            Write-Info "  • Run commands: Invoke-Command -Session `$global:RemoteSession -ScriptBlock { ... }"
            Write-Info "  • Copy files to remote: Copy-Item -Path 'local_file' -Destination 'remote_path' -ToSession `$global:RemoteSession"
            Write-Info "  • Copy files from remote: Copy-Item -Path 'remote_file' -Destination 'local_path' -FromSession `$global:RemoteSession"
            Write-Info "  • Close session: Remove-PSSession `$global:RemoteSession"
            Write-Host ""
            Write-Info "Example - Run your installation script remotely:"
            Write-Host "  Copy-Item -Path '.\Install-NtnxV4ApiEnvironment.ps1' -Destination 'C:\temp\' -ToSession `$global:RemoteSession" -ForegroundColor Gray
            Write-Host "  Invoke-Command -Session `$global:RemoteSession -ScriptBlock { C:\temp\Install-NtnxV4ApiEnvironment.ps1 }" -ForegroundColor Gray
        }
    }
    else {
        Write-Error "Session creation returned null - unknown error"
    }
}
catch {
    Write-Error "Failed to create PowerShell session: $($_.Exception.Message)"
    
    Write-Host ""
    Write-Warning "Troubleshooting steps:"
    Write-Info "1. Ensure PowerShell Remoting is enabled on target computer:"
    Write-Host "   Enable-PSRemoting -Force" -ForegroundColor Gray
    Write-Info "2. Check if WinRM service is running:"
    Write-Host "   Get-Service WinRM" -ForegroundColor Gray
    Write-Info "3. Configure trusted hosts if computers are not domain-joined:"
    Write-Host "   Set-Item WSMan:\localhost\Client\TrustedHosts -Value '$ComputerName' -Force" -ForegroundColor Gray
    Write-Info "4. Check firewall rules:"
    Write-Host "   Get-NetFirewallRule -Name '*WinRM*' | Where-Object Enabled -eq 'True'" -ForegroundColor Gray
    Write-Info "5. For HTTPS, ensure certificate is properly configured"
}

#endregion

#region Cleanup Function

# Register cleanup function
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    if ($global:RemoteSession -and $global:RemoteSession.State -eq 'Opened') {
        Write-Host "Cleaning up remote session..." -ForegroundColor Yellow
        Remove-PSSession $global:RemoteSession -ErrorAction SilentlyContinue
    }
} | Out-Null

#endregion

Write-Host ""
Write-Success "Remote session script completed!"