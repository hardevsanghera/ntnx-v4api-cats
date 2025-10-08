#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Enable PowerShell Remoting on Windows Computer
    
.DESCRIPTION
    This script enables PowerShell Remoting with optimal security settings
    Run this script as Administrator on the target computer
    
.PARAMETER AllowUnencrypted
    Allow unencrypted traffic (less secure, only for testing)
    
.PARAMETER TrustedHosts
    Set trusted hosts (use "*" for any host, or specific IPs/names)
    
.PARAMETER EnableCredSSP
    Enable CredSSP authentication (required for some scenarios)
    
.EXAMPLE
    .\Enable-PSRemoting.ps1
    
.EXAMPLE
    .\Enable-PSRemoting.ps1 -TrustedHosts "10.38.20.*" -EnableCredSSP
    
.NOTES
    - Requires Administrator privileges
    - Configures Windows Firewall automatically
    - Sets up WinRM service for remote management
#>

[CmdletBinding()]
param(
    [switch]$AllowUnencrypted,
    [string]$TrustedHosts = "",
    [switch]$EnableCredSSP
)

# Check if running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator"
    exit 1
}

Write-Host "🔧 Enabling PowerShell Remoting" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan

try {
    # Enable PowerShell Remoting
    Write-Host "📦 Enabling PowerShell Remoting..." -ForegroundColor Yellow
    Enable-PSRemoting -Force -SkipNetworkProfileCheck
    Write-Host "✅ PowerShell Remoting enabled" -ForegroundColor Green

    # Configure WinRM Service
    Write-Host "📦 Configuring WinRM Service..." -ForegroundColor Yellow
    Set-Service WinRM -StartupType Automatic
    Start-Service WinRM
    Write-Host "✅ WinRM Service configured" -ForegroundColor Green

    # Configure Windows Firewall
    Write-Host "📦 Configuring Windows Firewall..." -ForegroundColor Yellow
    Enable-NetFirewallRule -DisplayName "Windows Remote Management (HTTP-In)"
    Enable-NetFirewallRule -DisplayName "Windows Remote Management (HTTPS-In)"
    Write-Host "✅ Firewall rules enabled" -ForegroundColor Green

    # Set trusted hosts if specified
    if ($TrustedHosts) {
        Write-Host "📦 Setting trusted hosts to: $TrustedHosts" -ForegroundColor Yellow
        Set-Item WSMan:\localhost\Client\TrustedHosts -Value $TrustedHosts -Force
        Write-Host "✅ Trusted hosts configured" -ForegroundColor Green
    }

    # Enable CredSSP if requested
    if ($EnableCredSSP) {
        Write-Host "📦 Enabling CredSSP authentication..." -ForegroundColor Yellow
        Enable-WSManCredSSP -Role Server -Force
        Write-Host "✅ CredSSP enabled" -ForegroundColor Green
    }

    # Configure LocalAccountTokenFilterPolicy for local admin accounts
    Write-Host "📦 Configuring LocalAccountTokenFilterPolicy..." -ForegroundColor Yellow
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "LocalAccountTokenFilterPolicy" -Value 1 -Type DWord -Force
    Write-Host "✅ LocalAccountTokenFilterPolicy configured" -ForegroundColor Green

    # Allow unencrypted traffic if requested (NOT RECOMMENDED for production)
    if ($AllowUnencrypted) {
        Write-Host "⚠️  Enabling unencrypted traffic (NOT RECOMMENDED)" -ForegroundColor Yellow
        Set-Item WSMan:\localhost\Service\AllowUnencrypted -Value $true -Force
        Write-Host "✅ Unencrypted traffic allowed" -ForegroundColor Green
    }

    # Configure authentication methods
    Write-Host "📦 Configuring authentication methods..." -ForegroundColor Yellow
    Set-Item WSMan:\localhost\Service\Auth\Basic -Value $true -Force
    Set-Item WSMan:\localhost\Service\Auth\Negotiate -Value $true -Force
    Set-Item WSMan:\localhost\Service\Auth\Kerberos -Value $true -Force
    Write-Host "✅ Authentication methods configured" -ForegroundColor Green

    # Restart WinRM to apply all changes
    Write-Host "📦 Restarting WinRM service..." -ForegroundColor Yellow
    Restart-Service WinRM -Force
    Write-Host "✅ WinRM service restarted" -ForegroundColor Green

    # Test the configuration
    Write-Host "📦 Testing configuration..." -ForegroundColor Yellow
    $listeners = Get-WSManInstance -ResourceURI winrm/config/listener -SelectorSet @{Address="*";Transport="HTTP"}
    if ($listeners) {
        Write-Host "✅ HTTP listener is active" -ForegroundColor Green
    }

    # Display current configuration
    Write-Host ""
    Write-Host "📋 Current WinRM Configuration:" -ForegroundColor Cyan
    Write-Host "Service Status: $((Get-Service WinRM).Status)" -ForegroundColor White
    Write-Host "Startup Type: $((Get-Service WinRM).StartType)" -ForegroundColor White
    
    $config = winrm get winrm/config
    Write-Host "WinRM Configuration:" -ForegroundColor White
    $config | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }

    Write-Host ""
    Write-Host "🎉 PowerShell Remoting is now enabled!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📋 Test from remote computer:" -ForegroundColor Cyan
    Write-Host "  Test-WSMan -ComputerName $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "  Enter-PSSession -ComputerName $env:COMPUTERNAME -Credential (Get-Credential)" -ForegroundColor Gray

}
catch {
    Write-Error "Failed to enable PowerShell Remoting: $($_.Exception.Message)"
    exit 1
}