#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Copy PowerShell script to a remote Windows server
    
.DESCRIPTION
    This script copies the Install-NtnxV4ApiEnvironment.ps1 script (or any specified script)
    to a remote Windows computer and optionally executes it remotely.
    
.PARAMETER ComputerName
    The name or IP address of the remote computer
    Default: 10.38.20.187
    
.PARAMETER ScriptPath
    Local path to the script to copy
    Default: .\Install-NtnxV4ApiEnvironment.ps1
    
.PARAMETER RemoteDestination
    Remote destination path for the script
    Default: C:\temp\
    
.PARAMETER Credential
    PSCredential object for authentication. If not provided, will prompt for credentials
    
.PARAMETER ExecuteRemotely
    Execute the script on the remote computer after copying
    
.PARAMETER ScriptParameters
    Parameters to pass to the remote script (as hashtable)
    
.PARAMETER UseSSL
    Use HTTPS/SSL for the connection
    
.PARAMETER TestConnection
    Test the connection before attempting to copy
    
.EXAMPLE
    .\Copy-ScriptToRemote.ps1
    
.EXAMPLE
    .\Copy-ScriptToRemote.ps1 -ExecuteRemotely -TestConnection
    
.EXAMPLE
    .\Copy-ScriptToRemote.ps1 -ComputerName "server01" -ScriptParameters @{RepositoryPath="C:\NutanixAPI"; Force=$true}
    
.NOTES
    - Requires PowerShell Remoting to be enabled on target computer
    - May require administrator privileges for script execution
    - Creates destination directory if it doesn't exist
#>

[CmdletBinding()]
param(
    [string]$ComputerName = "10.38.20.154",
    [string]$ScriptPath = ".\Install-NtnxV4ApiEnvironment.ps1",
    [string]$RemoteDestination = "C:\temp\",
    [PSCredential]$Credential,
    [switch]$ExecuteRemotely,
    [hashtable]$ScriptParameters = @{},
    [switch]$UseSSL,
    [switch]$TestConnection
)

# Script configuration
$ErrorActionPreference = 'Stop'

Write-Host "📁 PowerShell Script Copy & Execute Tool" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
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
        [PSCredential]$Credential,
        [bool]$UseSSL
    )
    
    $port = if ($UseSSL) { 5986 } else { 5985 }
    
    Write-Info "Testing connection to $ComputerName on port $port..."
    
    # Test basic connectivity
    try {
        $tcpTest = Test-NetConnection -ComputerName $ComputerName -Port $port -WarningAction SilentlyContinue
        if (-not $tcpTest.TcpTestSucceeded) {
            throw "Port $port is not accessible"
        }
        Write-Success "Network connectivity test passed"
    }
    catch {
        Write-Warning "Network connectivity test failed: $($_.Exception.Message)"
        return $false
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
        Write-Success "PowerShell Remoting test passed"
        return $true
    }
    catch {
        Write-Warning "PowerShell Remoting test failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-ScriptSize {
    param([string]$Path)
    
    if (Test-Path $Path) {
        $size = (Get-Item $Path).Length
        if ($size -lt 1KB) {
            return "$size bytes"
        }
        elseif ($size -lt 1MB) {
            return "{0:N1} KB" -f ($size / 1KB)
        }
        else {
            return "{0:N1} MB" -f ($size / 1MB)
        }
    }
    return "Unknown"
}

#endregion

#region Validation

Write-Section "Pre-Copy Validation"

# Check if local script exists
if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script not found at path: $ScriptPath"
}

$scriptFile = Get-Item $ScriptPath
Write-Success "Found script: $($scriptFile.Name)"
Write-Info "Script size: $(Get-ScriptSize $ScriptPath)"
Write-Info "Last modified: $($scriptFile.LastWriteTime)"

# Validate destination path format
if (-not $RemoteDestination.EndsWith('\')) {
    $RemoteDestination += '\'
}

Write-Info "Target computer: $ComputerName"
Write-Info "Remote destination: $RemoteDestination"
Write-Info "Connection method: $(if ($UseSSL) { 'HTTPS/SSL' } else { 'HTTP' })"

#endregion

#region Credential Management

Write-Section "Authentication"

if (-not $Credential) {
    Write-Info "Prompting for credentials..."
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
    
    $connectionOk = Test-RemoteConnection -ComputerName $ComputerName -Credential $Credential -UseSSL:$UseSSL
    
    if (-not $connectionOk) {
        Write-Warning "Connection tests failed, but continuing anyway..."
        Write-Info "If copy fails, ensure PowerShell Remoting is enabled on the target computer"
    }
}

#endregion

#region Session Creation

Write-Section "Creating Remote Session"

try {
    $sessionParams = @{
        ComputerName = $ComputerName
        Credential = $Credential
        ErrorAction = 'Stop'
    }
    
    if ($UseSSL) {
        $sessionParams['UseSSL'] = $true
    }
    
    Write-Info "Creating PowerShell session..."
    $session = New-PSSession @sessionParams
    
    if ($session) {
        Write-Success "Remote session created successfully"
        Write-Info "Session ID: $($session.Id)"
        Write-Info "Computer Name: $($session.ComputerName)"
        
        # Test session with basic command
        $remoteInfo = Invoke-Command -Session $session -ScriptBlock {
            [PSCustomObject]@{
                ComputerName = $env:COMPUTERNAME
                PowerShellVersion = $PSVersionTable.PSVersion.ToString()
                FreeSpace = [math]::Round((Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace / 1GB, 2)
            }
        }
        
        Write-Info "Remote computer: $($remoteInfo.ComputerName)"
        Write-Info "PowerShell version: $($remoteInfo.PowerShellVersion)"
        Write-Info "C: drive free space: $($remoteInfo.FreeSpace) GB"
    }
}
catch {
    Write-Error "Failed to create remote session: $($_.Exception.Message)"
}

#endregion

#region Directory Preparation

Write-Section "Preparing Remote Directory"

try {
    Write-Info "Ensuring remote destination directory exists..."
    
    Invoke-Command -Session $session -ScriptBlock {
        param($DestPath)
        
        if (-not (Test-Path $DestPath)) {
            New-Item -Path $DestPath -ItemType Directory -Force | Out-Null
            Write-Output "Created directory: $DestPath"
        }
        else {
            Write-Output "Directory already exists: $DestPath"
        }
        
        # Check write permissions
        $testFile = Join-Path $DestPath "write_test.tmp"
        try {
            "test" | Out-File -FilePath $testFile -ErrorAction Stop
            Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            Write-Output "Write permissions confirmed"
        }
        catch {
            throw "No write permissions to $DestPath"
        }
    } -ArgumentList $RemoteDestination
    
    Write-Success "Remote directory is ready"
}
catch {
    Write-Error "Failed to prepare remote directory: $($_.Exception.Message)"
}

#endregion

#region File Copy

Write-Section "Copying Script to Remote Computer"

try {
    $remoteScriptPath = Join-Path $RemoteDestination $scriptFile.Name
    
    Write-Info "Copying $($scriptFile.Name) to $ComputerName..."
    Write-Info "Source: $ScriptPath"
    Write-Info "Destination: $remoteScriptPath"
    
    $copyStartTime = Get-Date
    Copy-Item -Path $ScriptPath -Destination $RemoteDestination -ToSession $session -Force
    $copyDuration = (Get-Date) - $copyStartTime
    
    Write-Success "File copied successfully in $([math]::Round($copyDuration.TotalSeconds, 2)) seconds"
    
    # Verify the copy
    $remoteFileInfo = Invoke-Command -Session $session -ScriptBlock {
        param($FilePath)
        if (Test-Path $FilePath) {
            $file = Get-Item $FilePath
            [PSCustomObject]@{
                Exists = $true
                Size = $file.Length
                LastWriteTime = $file.LastWriteTime
            }
        }
        else {
            [PSCustomObject]@{
                Exists = $false
            }
        }
    } -ArgumentList $remoteScriptPath
    
    if ($remoteFileInfo.Exists) {
        Write-Success "Copy verification successful"
        Write-Info "Remote file size: $(Get-ScriptSize -Path $ScriptPath)"
        Write-Info "Remote file time: $($remoteFileInfo.LastWriteTime)"
    }
    else {
        Write-Error "Copy verification failed - file not found on remote computer"
    }
}
catch {
    Write-Error "Failed to copy script: $($_.Exception.Message)"
}

#endregion

#region Remote Execution - COMMENTED OUT

<# 
REMOTE EXECUTION DISABLED - Script will only be copied, not executed
To enable remote execution, uncomment this section and use -ExecuteRemotely parameter

if ($ExecuteRemotely) {
    Write-Section "Executing Script Remotely"
    
    try {
        Write-Info "Preparing to execute script on remote computer..."
        
        # Build parameter string for script execution
        $paramString = ""
        if ($ScriptParameters.Count -gt 0) {
            $paramArray = @()
            foreach ($param in $ScriptParameters.GetEnumerator()) {
                if ($param.Value -is [bool]) {
                    if ($param.Value) {
                        $paramArray += "-$($param.Key)"
                    }
                }
                else {
                    $paramArray += "-$($param.Key) `"$($param.Value)`""
                }
            }
            $paramString = $paramArray -join " "
            Write-Info "Script parameters: $paramString"
        }
        
        Write-Info "Executing: $remoteScriptPath $paramString"
        Write-Warning "Script execution may take several minutes..."
        
        $executionStartTime = Get-Date
        
        $executionResult = Invoke-Command -Session $session -ScriptBlock {
            param($ScriptPath, $Parameters)
            
            # Change to script directory
            $scriptDir = Split-Path $ScriptPath -Parent
            Set-Location $scriptDir
            
            # Execute script
            if ($Parameters) {
                $command = "& `"$ScriptPath`" $Parameters"
                Invoke-Expression $command
            }
            else {
                & $ScriptPath
            }
        } -ArgumentList $remoteScriptPath, $paramString
        
        $executionDuration = (Get-Date) - $executionStartTime
        
        Write-Success "Script execution completed in $([math]::Round($executionDuration.TotalMinutes, 2)) minutes"
        
        if ($executionResult) {
            Write-Info "Execution output:"
            $executionResult | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        }
    }
    catch {
        Write-Warning "Script execution encountered an error: $($_.Exception.Message)"
        Write-Info "The script was copied successfully but execution failed"
        Write-Info "You can execute it manually by connecting to the remote computer"
    }
}
#>

#endregion

#region Summary and Cleanup

Write-Section "Summary"

Write-Success "Operation completed successfully!"
Write-Info "Script location on remote computer: $remoteScriptPath"

if (-not $ExecuteRemotely) {
    Write-Info ""
    Write-Info "To execute the script manually:"
    Write-Host "  1. Connect to remote computer: Enter-PSSession -ComputerName $ComputerName -Credential `$Credential" -ForegroundColor Gray
    Write-Host "  2. Navigate to directory: cd $RemoteDestination" -ForegroundColor Gray
    Write-Host "  3. Execute script: .\$($scriptFile.Name)" -ForegroundColor Gray
    Write-Info ""
    Write-Info "Or execute remotely:"
    Write-Host "  Invoke-Command -ComputerName $ComputerName -Credential `$Credential -ScriptBlock { $remoteScriptPath }" -ForegroundColor Gray
}

# Cleanup
if ($session) {
    Write-Info "Cleaning up remote session..."
    Remove-PSSession $session
    Write-Success "Remote session closed"
}

Write-Host ""
Write-Success "Script copy operation completed!"

#endregion