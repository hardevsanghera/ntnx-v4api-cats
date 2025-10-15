 #!/usr/bin/env pwsh
<#
.SYNOPSIS
    Automated installation script for ntnx-v4api-cats environment setup
    Download just this script from the repository:
    cd to your directory then
    curl -O https://raw.githubusercontent.com/hardevsanghera/ntnx-v4api-cats/main/experimental/Install-NtnxV4ApiEnvironment.ps1
 
.DESCRIPTION
    --------------------------------------
    -EXPERIMENTAL - Use at your own risk!-
    --------------------------------------
    This PowerShell 7 script automatically installs all necessary components for the 
    ntnx-v4api-cats repository to run successfully on Windows. It installs:
    - PowerShell 7 (if not already installed)
    - Python 3.13+ (if not already installed)  
    - Visual Studio Code (if not already installed)
    - Git for Windows (if not already installed)
    - Required PowerShell modules (ImportExcel)
    - Python virtual environment with dependencies from requirements.txt
    
    Based on the repository: https://github.com/hardevsanghera/ntnx-v4api-cats
    
.PARAMETER RepositoryPath
    Local path where the ntnx-v4api-cats repository will be cloned or already exists
    Default: $env:USERPROFILE\Documents\ntnx-v4api-cats
    
.PARAMETER SkipGitClone
    If specified, skips cloning the repository (assumes it already exists at RepositoryPath)
    
.PARAMETER Force
    Forces reinstallation of components even if they appear to be already installed
    
.EXAMPLE
    .\Install-NtnxV4ApiEnvironment.ps1
    
.EXAMPLE
    .\Install-NtnxV4ApiEnvironment.ps1 -RepositoryPath "C:\Dev\ntnx-v4api-cats" -SkipGitClone
    
.NOTES
    - Requires administrator privileges for some installations
    - Based on repository requirements: PowerShell 7, Python 3.13+, VS Code, Git for Windows
    - Creates Python virtual environment in repository\.venv
    - Installs dependencies from files\requirements.txt
#>

[CmdletBinding()]
param(
    [string]$RepositoryPath = "$env:USERPROFILE\Documents\ntnx-v4api-cats",
    [switch]$SkipGitClone,
    [switch]$Force
)

# Script configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Repository URL
$RepoUrl = "https://github.com/hardevsanghera/ntnx-v4api-cats.git"

# Expected Python packages based on repository analysis
$PythonRequirements = @"
requests>=2.31.0
pandas>=2.0.0
openpyxl>=3.1.0
urllib3>=2.0.0
"@

Write-Host "  Nutanix v4 API Environment Setup Script" -ForegroundColor Cyan
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

#region Helper Functions

function Test-IsAdmin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host " $Title" -ForegroundColor Yellow
    Write-Host ("-" * ($Title.Length + 3)) -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Message)
    Write-Host " $Message" -ForegroundColor Green
}

function Write-Info {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Green
}

function Write-Warning {
    param([string]$Message)
    Write-Host "  $Message" -ForegroundColor Yellow
}

function Test-CommandExists {
    param([string]$Command)
    try {
        Get-Command $Command -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Get-LatestGitHubRelease {
    param([string]$Repository)
    try {
        $apiUrl = "https://api.github.com/repos/$Repository/releases/latest"
        $response = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
        return $response.tag_name
    }
    catch {
        Write-Warning "Could not fetch latest version for $Repository"
        return $null
    }
}

function Install-FromUrl {
    param(
        [string]$Url,
        [string]$OutputPath,
        [string]$Arguments = "",
        [string]$Description
    )
    
    Write-Info "Downloading $Description..."
    try {
        Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
        Write-Info "Installing $Description..."
        
        if ($Arguments) {
            Start-Process -FilePath $OutputPath -ArgumentList $Arguments -Wait -NoNewWindow
        }
        else {
            Start-Process -FilePath $OutputPath -Wait -NoNewWindow
        }
        
        Write-Success "$Description installed successfully"
        
        # Clean up installer
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Error "Failed to install $Description`: $($_.Exception.Message)"
    }
}

#endregion

#region Pre-flight Checks

Write-Section "Pre-flight Checks"

# Check if running on Windows
if ($PSVersionTable.Platform -and $PSVersionTable.Platform -ne "Win32NT") {
    Write-Error "This script is designed for Windows only"
}

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion
Write-Info "Current PowerShell version: $psVersion"

if ($psVersion.Major -lt 7) {
    Write-Warning "PowerShell 7+ is required. Current version is $psVersion"
    $needsPwsh = $true
}
else {
    Write-Success "PowerShell 7+ is already installed"
    $needsPwsh = $false
}

# Check admin privileges for installations
if (-not (Test-IsAdmin)) {
    Write-Warning "Some installations may require administrator privileges"
    Write-Info "Consider running as administrator if installations fail"
}

#endregion

#region PowerShell 7 Installation

if ($needsPwsh -or $Force) {
    Write-Section "Installing PowerShell 7"
    
    try {
        # Use winget if available, otherwise download directly
        if (Test-CommandExists "winget") {
            Write-Info "Installing PowerShell 7 via winget..."
            try {
                winget install --id Microsoft.Powershell --source winget --silent --accept-package-agreements --accept-source-agreements
                Write-Success "PowerShell 7 installed via winget"
            }
            catch {
                Write-Warning "Winget installation failed, trying direct download..."
                $useDirectDownload = $true
            }
        }
        else {
            $useDirectDownload = $true
        }
        
        if ($useDirectDownload) {
            # Download and install PowerShell 7 manually with better error handling
            Write-Info "Attempting direct download of PowerShell 7..."
            
            # Try multiple download sources
            $downloadAttempts = @(
                @{
                    Url = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.6/PowerShell-7.4.6-win-x64.msi"
                    Name = "PowerShell 7.4.6"
                },
                @{
                    Url = "https://github.com/PowerShell/PowerShell/releases/download/v7.4.5/PowerShell-7.4.5-win-x64.msi"
                    Name = "PowerShell 7.4.5 (fallback)"
                }
            )
            
            $installSuccess = $false
            foreach ($attempt in $downloadAttempts) {
                try {
                    $installerPath = "$env:TEMP\PowerShell-installer.msi"
                    
                    # Clean up any existing installer
                    if (Test-Path $installerPath) {
                        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                    }
                    
                    Write-Info "Downloading $($attempt.Name)..."
                    try {
                        Invoke-WebRequest -Uri $attempt.Url -OutFile $installerPath -UseBasicParsing -TimeoutSec 300
                    }
                    catch {
                        Write-Warning "Download failed: $($_.Exception.Message)"
                        continue
                    }
                    
                    # Verify the download
                    if ((Test-Path $installerPath) -and ((Get-Item $installerPath).Length -gt 1MB)) {
                        Write-Info "Installing $($attempt.Name)..."
                        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installerPath`" /quiet /norestart" -Wait -PassThru -NoNewWindow
                        
                        if ($process.ExitCode -eq 0) {
                            Write-Success "$($attempt.Name) installed successfully"
                            $installSuccess = $true
                            break
                        }
                        else {
                            Write-Warning "$($attempt.Name) installation failed with exit code: $($process.ExitCode)"
                        }
                    }
                    else {
                        Write-Warning "Downloaded file appears to be invalid or incomplete"
                    }
                }
                catch {
                    Write-Warning "Failed to download/install $($attempt.Name): $($_.Exception.Message)"
                }
                finally {
                    # Clean up installer
                    if (Test-Path $installerPath) {
                        Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            
            if (-not $installSuccess) {
                Write-Warning "PowerShell 7 installation failed. You may need to install it manually from:"
                Write-Warning "https://github.com/PowerShell/PowerShell/releases/latest"
                Write-Info "The script will continue with the current PowerShell version"
            }
        }
        
        Write-Info "You may need to restart your terminal to use the new PowerShell version"
    }
    catch {
        Write-Warning "PowerShell 7 installation encountered an error: $($_.Exception.Message)"
        Write-Info "Continuing with current PowerShell version..."
    }
}

#endregion

#region Python 3.13+ Installation

Write-Section "Checking Python Installation"

$pythonInstalled = $false
$pythonVersion = $null

# Check for Python
if (Test-CommandExists "python") {
    try {
        $pythonVersionOutput = python --version 2>&1
        if ($pythonVersionOutput -match "Python (\d+\.\d+\.\d+)") {
            $pythonVersion = [Version]$matches[1]
            $pythonInstalled = $true
            Write-Info "Found Python version: $pythonVersion"
            
            if ($pythonVersion -ge [Version]"3.13.0") {
                Write-Success "Python 3.13+ is already installed"
                $needsPython = $false
            }
            else {
                Write-Warning "Python version $pythonVersion is below required 3.13+"
                $needsPython = $true
            }
        }
    }
    catch {
        Write-Warning "Could not determine Python version"
        $needsPython = $true
    }
}
else {
    Write-Info "Python not found in PATH"
    $needsPython = $true
}

if ($needsPython -or $Force) {
    Write-Section "Installing Python 3.13+"
    
    try {
        # Use winget if available
        if (Test-CommandExists "winget") {
            Write-Info "Installing Python 3.13 via winget..."
            try {
                winget install --id Python.Python.3.13 --source winget --silent --accept-package-agreements --accept-source-agreements
                Write-Success "Python 3.13 installed via winget"
            }
            catch {
                Write-Warning "Winget installation failed, trying direct download..."
                $useDirectDownload = $true
            }
        }
        else {
            $useDirectDownload = $true
        }
        
        if ($useDirectDownload) {
            # Download Python directly from python.org with better error handling
            Write-Info "Attempting direct download of Python 3.13..."
            $pythonUrl = "https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe"
            $pythonInstaller = "$env:TEMP\python-3.13.0-amd64.exe"
            
            try {
                # Clean up any existing installer
                if (Test-Path $pythonInstaller) {
                    Remove-Item $pythonInstaller -Force -ErrorAction SilentlyContinue
                }
                
                Write-Info "Downloading Python 3.13..."
                try {
                    Invoke-WebRequest -Uri $pythonUrl -OutFile $pythonInstaller -UseBasicParsing -TimeoutSec 300
                }
                catch {
                    Write-Warning "Python download failed: $($_.Exception.Message)"
                    return
                }
                
                # Verify the download
                if ((Test-Path $pythonInstaller) -and ((Get-Item $pythonInstaller).Length -gt 10MB)) {
                    Write-Info "Installing Python 3.13..."
                    $pythonArgs = "/quiet InstallAllUsers=1 PrependPath=1 Include_test=0"
                    $process = Start-Process -FilePath $pythonInstaller -ArgumentList $pythonArgs -Wait -PassThru -NoNewWindow
                    
                    if ($process.ExitCode -eq 0) {
                        Write-Success "Python 3.13 installed successfully"
                    }
                    else {
                        Write-Warning "Python installation failed with exit code: $($process.ExitCode)"
                    }
                }
                else {
                    Write-Warning "Downloaded Python installer appears to be invalid or incomplete"
                }
            }
            catch {
                Write-Warning "Failed to download/install Python: $($_.Exception.Message)"
            }
            finally {
                # Clean up installer
                if (Test-Path $pythonInstaller) {
                    Remove-Item $pythonInstaller -Force -ErrorAction SilentlyContinue
                }
            }
        }
        
        Write-Info "Refreshing environment variables..."
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Verify Python installation
        Start-Sleep -Seconds 3
        if (Test-CommandExists "python") {
            $newPythonVersion = python --version 2>&1
            Write-Success "Python installation verified: $newPythonVersion"
        }
        else {
            Write-Warning "Python installation may require a system restart to be fully functional"
        }
    }
    catch {
        Write-Warning "Python installation encountered an error: $($_.Exception.Message)"
        Write-Info "You may need to install Python manually from: https://www.python.org/downloads/"
    }
}

#endregion

#region Visual Studio Code Installation

Write-Section "Checking Visual Studio Code Installation"

if (-not (Test-CommandExists "code") -or $Force) {
    Write-Section "Installing Visual Studio Code"
    
    try {
        # Use winget if available
        if (Test-CommandExists "winget") {
            Write-Info "Installing Visual Studio Code via winget..."
            try {
                winget install --id Microsoft.VisualStudioCode --source winget --silent --accept-package-agreements --accept-source-agreements
                Write-Success "Visual Studio Code installed via winget"
            }
            catch {
                Write-Warning "Winget installation failed, trying direct download..."
                $useDirectDownload = $true
            }
        }
        else {
            $useDirectDownload = $true
        }
        
        if ($useDirectDownload) {
            # Download VS Code directly with better error handling
            Write-Info "Attempting direct download of Visual Studio Code..."
            
            # Try multiple download sources
            $downloadAttempts = @(
                @{
                    Url = "https://code.visualstudio.com/sha/download?build=stable&os=win32-x64-user"
                    Name = "VS Code (User Installer)"
                },
                @{
                    Url = "https://update.code.visualstudio.com/latest/win32-x64/stable"
                    Name = "VS Code (System Installer - fallback)"
                }
            )
            
            $installSuccess = $false
            foreach ($attempt in $downloadAttempts) {
                try {
                    $vscodeInstaller = "$env:TEMP\VSCodeSetup.exe"
                    
                    # Clean up any existing installer
                    if (Test-Path $vscodeInstaller) {
                        Remove-Item $vscodeInstaller -Force -ErrorAction SilentlyContinue
                    }
                    
                    Write-Info "Downloading $($attempt.Name)..."
                    try {
                        Invoke-WebRequest -Uri $attempt.Url -OutFile $vscodeInstaller -UseBasicParsing -TimeoutSec 300
                    }
                    catch {
                        Write-Warning "Download failed: $($_.Exception.Message)"
                        continue
                    }
                    
                    # Verify the download
                    if ((Test-Path $vscodeInstaller) -and ((Get-Item $vscodeInstaller).Length -gt 50MB)) {
                        Write-Info "Installing $($attempt.Name)..."
                        $vscodeArgs = "/SILENT /mergetasks=!runcode,addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath"
                        $process = Start-Process -FilePath $vscodeInstaller -ArgumentList $vscodeArgs -Wait -PassThru -NoNewWindow
                        
                        if ($process.ExitCode -eq 0) {
                            Write-Success "$($attempt.Name) installed successfully"
                            $installSuccess = $true
                            break
                        }
                        else {
                            Write-Warning "$($attempt.Name) installation failed with exit code: $($process.ExitCode)"
                        }
                    }
                    else {
                        Write-Warning "Downloaded VS Code installer appears to be invalid or incomplete"
                    }
                }
                catch {
                    Write-Warning "Failed to download/install $($attempt.Name): $($_.Exception.Message)"
                }
                finally {
                    # Clean up installer
                    if (Test-Path $vscodeInstaller) {
                        Remove-Item $vscodeInstaller -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            
            if (-not $installSuccess) {
                Write-Warning "Visual Studio Code installation failed. You may need to install it manually from:"
                Write-Warning "https://code.visualstudio.com/download"
                Write-Info "The script will continue without VS Code"
            }
        }
        
        Write-Info "Refreshing environment variables..."
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Verify VS Code installation
        Start-Sleep -Seconds 3
        if (Test-CommandExists "code") {
            $vscodeVersion = code --version 2>&1 | Select-Object -First 1
            Write-Success "Visual Studio Code installation verified: $vscodeVersion"
        }
        else {
            Write-Warning "VS Code installation may require a system restart to be fully functional"
        }
    }
    catch {
        Write-Warning "Visual Studio Code installation encountered an error: $($_.Exception.Message)"
        Write-Info "You may need to install VS Code manually from: https://code.visualstudio.com/download"
    }
}
else {
    $vscodeVersion = code --version 2>&1 | Select-Object -First 1
    Write-Success "Visual Studio Code is already installed: $vscodeVersion"
}

#endregion

#region Git for Windows Installation

Write-Section "Checking Git Installation"

if (-not (Test-CommandExists "git") -or $Force) {
    Write-Section "Installing Git for Windows"
    
    try {
        # Use winget if available
        if (Test-CommandExists "winget") {
            Write-Info "Installing Git for Windows via winget..."
            try {
                winget install --id Git.Git --source winget --silent --accept-package-agreements --accept-source-agreements
                Write-Success "Git for Windows installed via winget"
            }
            catch {
                Write-Warning "Winget installation failed, trying direct download..."
                $useDirectDownload = $true
            }
        }
        else {
            $useDirectDownload = $true
        }
        
        if ($useDirectDownload) {
            # Download Git for Windows directly with better error handling
            Write-Info "Attempting direct download of Git for Windows..."
            
            # Try multiple download sources
            $downloadAttempts = @(
                @{
                    Url = "https://github.com/git-for-windows/git/releases/download/v2.51.0.windows.2/Git-2.51.0.2-64-bit.exe"
                    Name = "Git 2.51.0"
                },
                @{
                    Url = "https://github.com/git-for-windows/git/releases/download/v2.47.0.windows.2/Git-2.47.0.2-64-bit.exe"
                    Name = "Git 2.47.0 (fallback)"
                }
            )
            
            $installSuccess = $false
            foreach ($attempt in $downloadAttempts) {
                try {
                    $gitInstaller = "$env:TEMP\Git-installer.exe"
                    
                    # Clean up any existing installer
                    if (Test-Path $gitInstaller) {
                        Remove-Item $gitInstaller -Force -ErrorAction SilentlyContinue
                    }
                    
                    Write-Info "Downloading $($attempt.Name)..."
                    try {
                        Invoke-WebRequest -Uri $attempt.Url -OutFile $gitInstaller -UseBasicParsing -TimeoutSec 300
                    }
                    catch {
                        Write-Warning "Download failed: $($_.Exception.Message)"
                        continue
                    }
                    
                    # Verify the download
                    if ((Test-Path $gitInstaller) -and ((Get-Item $gitInstaller).Length -gt 10MB)) {
                        Write-Info "Installing $($attempt.Name)..."
                        $gitArgs = "/SILENT /COMPONENTS=`"icons,ext\reg\shellhere,assoc,assoc_sh`""
                        $process = Start-Process -FilePath $gitInstaller -ArgumentList $gitArgs -Wait -PassThru -NoNewWindow
                        
                        if ($process.ExitCode -eq 0) {
                            Write-Success "$($attempt.Name) installed successfully"
                            $installSuccess = $true
                            break
                        }
                        else {
                            Write-Warning "$($attempt.Name) installation failed with exit code: $($process.ExitCode)"
                        }
                    }
                    else {
                        Write-Warning "Downloaded Git installer appears to be invalid or incomplete"
                    }
                }
                catch {
                    Write-Warning "Failed to download/install $($attempt.Name): $($_.Exception.Message)"
                }
                finally {
                    # Clean up installer
                    if (Test-Path $gitInstaller) {
                        Remove-Item $gitInstaller -Force -ErrorAction SilentlyContinue
                    }
                }
            }
            
            if (-not $installSuccess) {
                Write-Warning "Git for Windows installation failed. You may need to install it manually from:"
                Write-Warning "https://git-scm.com/download/win"
                Write-Info "The script will continue, but repository cloning may not work"
            }
        }
        
        Write-Info "Refreshing environment variables..."
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
        
        # Verify Git installation
        Start-Sleep -Seconds 3
        if (Test-CommandExists "git") {
            $gitVersion = git --version 2>&1
            Write-Success "Git installation verified: $gitVersion"
        }
        else {
            Write-Warning "Git installation may require a system restart to be fully functional"
        }
    }
    catch {
        Write-Warning "Git for Windows installation encountered an error: $($_.Exception.Message)"
        Write-Info "You may need to install Git manually from: https://git-scm.com/download/win"
    }
}
else {
    $gitVersion = git --version 2>&1
    Write-Success "Git is already installed: $gitVersion"
}

#endregion

#region Repository Setup

Write-Section "Repository Setup"

if (-not $SkipGitClone) {
    if (Test-Path $RepositoryPath) {
        if ($Force) {
            Write-Info "Removing existing repository directory..."
            Remove-Item $RepositoryPath -Recurse -Force
        }
        else {
            Write-Warning "Repository already exists at $RepositoryPath"
            Write-Info "Use -Force to overwrite or -SkipGitClone to use existing repository"
            $response = Read-Host "Continue with existing repository? (y/N)"
            if ($response -notmatch "^[Yy]") {
                Write-Error "Repository setup cancelled by user"
            }
            $SkipGitClone = $true
        }
    }
    
    if (-not $SkipGitClone) {
        Write-Info "Cloning repository to $RepositoryPath..."
        try {
            if (Test-CommandExists "git") {
                # Use cmd.exe to run git to avoid PowerShell stderr interpretation issues
                Write-Info "Cloning repository via Git..."
                
                # Create the parent directory if it doesn't exist
                $parentDir = Split-Path $RepositoryPath -Parent
                if (-not (Test-Path $parentDir)) {
                    New-Item -Path $parentDir -ItemType Directory -Force | Out-Null
                }
                
                # Use cmd.exe to execute git clone to avoid PowerShell stderr issues
                $gitCommand = "git clone `"$RepoUrl`" `"$RepositoryPath`""
                Write-Info "Executing: $gitCommand"
                
                # Execute via cmd.exe to properly handle git's output
                $gitProcess = Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $gitCommand -Wait -PassThru -NoNewWindow -RedirectStandardOutput "$env:TEMP\git_output.txt" -RedirectStandardError "$env:TEMP\git_error.txt"
                $gitExitCode = $gitProcess.ExitCode
                
                # Read the output files
                $gitOutput = @()
                if (Test-Path "$env:TEMP\git_output.txt") {
                    $gitOutput += Get-Content "$env:TEMP\git_output.txt" -ErrorAction SilentlyContinue
                    Remove-Item "$env:TEMP\git_output.txt" -Force -ErrorAction SilentlyContinue
                }
                if (Test-Path "$env:TEMP\git_error.txt") {
                    $gitOutput += Get-Content "$env:TEMP\git_error.txt" -ErrorAction SilentlyContinue
                    Remove-Item "$env:TEMP\git_error.txt" -Force -ErrorAction SilentlyContinue
                }
                
                if ($gitExitCode -eq 0 -and (Test-Path $RepositoryPath)) {
                    Write-Success "Repository cloned successfully"
                    if ($gitOutput) {
                        Write-Info "Git output: $($gitOutput -join "`n")"
                    }
                }
                else {
                    Write-Warning "Repository cloning failed (Exit code: $gitExitCode)"
                    if ($gitOutput) {
                        Write-Warning "Git output: $($gitOutput -join "`n")"
                    }
                    
                    # Common troubleshooting suggestions
                    Write-Info "Possible causes:"
                    Write-Info " Network connectivity issues"
                    Write-Info " Repository access permissions"
                    Write-Info " Directory already exists and is not empty"
                    Write-Info " Git credentials not configured"
                    Write-Info ""
                    Write-Info "You can clone manually with: git clone `"$RepoUrl`" `"$RepositoryPath`""
                    Write-Info "The script will continue with existing setup..."
                }
            }
            else {
                Write-Warning "Git is not available - cannot clone repository"
                Write-Info "Please clone manually: git clone `"$RepoUrl`" `"$RepositoryPath`""
                Write-Info "Or use -SkipGitClone if repository already exists"
            }
        }
        catch {
            Write-Warning "Failed to clone repository: $($_.Exception.Message)"
            Write-Info "You can clone manually with: git clone $RepoUrl `"$RepositoryPath`""
            Write-Info "The script will continue with existing setup..."
        }
    }
}

# Verify repository exists
if (-not (Test-Path $RepositoryPath)) {
    Write-Warning "Repository not found at $RepositoryPath"
    Write-Info "You can:"
    Write-Info "1. Clone manually: git clone $RepoUrl $RepositoryPath"
    Write-Info "2. Run script again after manual clone"
    Write-Info "3. Use -RepositoryPath to specify existing repository location"
    Write-Info ""
    Write-Info "The script will continue with component installation only..."
    $repositoryAvailable = $false
}
else {
    Write-Success "Repository available at: $RepositoryPath"
    $repositoryAvailable = $true
}

#endregion

#region VS Code Configuration

Write-Section "Configuring Visual Studio Code"

if ($repositoryAvailable -and (Test-CommandExists "code")) {
    try {
        Write-Info "Configuring VS Code settings for the project..."
        
        # Create .vscode directory in repository if it doesn't exist
        $vscodeDir = Join-Path $RepositoryPath ".vscode"
        if (-not (Test-Path $vscodeDir)) {
            New-Item -Path $vscodeDir -ItemType Directory -Force | Out-Null
            Write-Info "Created .vscode directory"
        }
        
        # Configure VS Code settings.json
        $settingsPath = Join-Path $vscodeDir "settings.json"
        
        # Detect installed paths
        $gitPath = $null
        $pythonPath = $null
        $pwshPath = $null
        
        # Find Git executable
        if (Test-CommandExists "git") {
            try {
                $gitPath = (Get-Command git).Source
                Write-Info "Found Git at: $gitPath"
            }
            catch {
                Write-Warning "Could not determine Git path"
            }
        }
        
        # Find Python executable (prefer the one in our venv)
        $venvPythonPath = Join-Path $RepositoryPath ".venv\Scripts\python.exe"
        if (Test-Path $venvPythonPath) {
            $pythonPath = $venvPythonPath
            Write-Info "Found Python (venv) at: $pythonPath"
        }
        elseif (Test-CommandExists "python") {
            try {
                $pythonPath = (Get-Command python).Source
                Write-Info "Found Python at: $pythonPath"
            }
            catch {
                Write-Warning "Could not determine Python path"
            }
        }
        
        # Find PowerShell 7 executable
        if (Test-CommandExists "pwsh") {
            try {
                $pwshPath = (Get-Command pwsh).Source
                Write-Info "Found PowerShell 7 at: $pwshPath"
            }
            catch {
                Write-Warning "Could not determine PowerShell 7 path"
            }
        }
        
        # Create VS Code settings
        $vscodeSettings = @{
            # Git configuration
            "git.enabled" = $true
            "git.autorefresh" = $true
            "git.autofetch" = $true
            
            # Terminal configuration - use PowerShell 7 as default
            "terminal.integrated.defaultProfile.windows" = "PowerShell Core"
            "terminal.integrated.profiles.windows" = @{
                "PowerShell Core" = @{
                    "source" = "PowerShell"
                    "icon" = "terminal-powershell"
                }
                "PowerShell" = @{
                    "source" = "PowerShell"
                    "icon" = "terminal-powershell"
                }
                "Command Prompt" = @{
                    "path" = @(
                        "${env:windir}\Sysnative\cmd.exe",
                        "${env:windir}\System32\cmd.exe"
                    )
                    "args" = @()
                    "icon" = "terminal-cmd"
                }
            }
            
            # Python configuration
            "python.defaultInterpreterPath" = if ($pythonPath) { $pythonPath.Replace('\', '/') } else { "python" }
            "python.terminal.activateEnvironment" = $true
            "python.terminal.activateEnvInCurrentTerminal" = $true
            
            # PowerShell configuration
            "powershell.powerShellDefaultVersion" = "PowerShell Core"
            
            # File associations
            "files.associations" = @{
                "*.ps1" = "powershell"
                "*.psm1" = "powershell"
                "*.psd1" = "powershell"
            }
            
            # Explorer settings
            "explorer.confirmDelete" = $false
            "explorer.confirmDragAndDrop" = $false
            "explorer.autoReveal" = "focusNoScroll"
            "explorer.autoRevealExclude" = @{
                "**/node_modules" = $true
                "**/.git" = $true
            }
            
            # Editor settings
            "editor.minimap.enabled" = $true
            "editor.wordWrap" = "on"
            "editor.renderWhitespace" = "boundary"
            
            # Workbench settings - disable welcome screen and auto-expand folders
            "workbench.startupEditor" = "none"
            "workbench.welcomePage.walkthroughs.openOnInstall" = $false
            "workbench.tips.enabled" = $false
            "explorer.openEditors.visible" = 0
            "workbench.tree.renderIndentGuides" = "always"
            "explorer.sortOrder" = "type"
            "explorer.compactFolders" = $false
            "explorer.expandSingleFolderWorkspaces" = $true
            
            # Security settings - auto-trust workspace
            "security.workspace.trust.enabled" = $false
            "security.workspace.trust.startupPrompt" = "never"
            "security.workspace.trust.banner" = "never"
            "security.workspace.trust.emptyWindow" = $false
        }
        
        # Add Git path if found
        if ($gitPath) {
            $vscodeSettings["git.path"] = $gitPath.Replace('\', '/')
        }
        
        # Add PowerShell 7 path if found
        if ($pwshPath) {
            $vscodeSettings["terminal.integrated.profiles.windows"]["PowerShell Core"]["path"] = $pwshPath.Replace('\', '/')
            $vscodeSettings["powershell.powerShellExePath"] = $pwshPath.Replace('\', '/')
        }
        
        # Convert to JSON and save
        $settingsJson = $vscodeSettings | ConvertTo-Json -Depth 5
        $settingsJson | Set-Content -Path $settingsPath -Encoding UTF8
        Write-Success "VS Code settings configured at: $settingsPath"
        
        # Create launch.json for Python debugging
        $launchPath = Join-Path $vscodeDir "launch.json"
        $launchConfig = @{
            "version" = "0.2.0"
            "configurations" = @(
                @{
                    "name" = "Python: Current File"
                    "type" = "python"
                    "request" = "launch"
                    "program" = "`${file}"
                    "console" = "integratedTerminal"
                    "cwd" = "`${workspaceFolder}"
                    "env" = @{
                        "PYTHONPATH" = "`${workspaceFolder}"
                    }
                },
                @{
                    "name" = "Python: update_categories_for_vm.py"
                    "type" = "python"
                    "request" = "launch"
                    "program" = "`${workspaceFolder}/update_categories_for_vm.py"
                    "console" = "integratedTerminal"
                    "cwd" = "`${workspaceFolder}"
                    "env" = @{
                        "PYTHONPATH" = "`${workspaceFolder}"
                    }
                }
            )
        }
        
        $launchJson = $launchConfig | ConvertTo-Json -Depth 5
        $launchJson | Set-Content -Path $launchPath -Encoding UTF8
        Write-Success "VS Code launch configuration created at: $launchPath"
        
        <# Commented out - VS Code tasks creation
        # Create tasks.json for PowerShell scripts
        $tasksPath = Join-Path $vscodeDir "tasks.json"
        $tasksConfig = @{
            "version" = "2.0.0"
            "tasks" = @(
                @{
                    "label" = "Run PowerShell Script"
                    "type" = "shell"
                    "command" = if ($pwshPath) { $pwshPath } else { "pwsh" }
                    "args" = @("-File", "`${file}")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                    "problemMatcher" = @()
                },
                @{
                    "label" = "List VMs"
                    "type" = "shell"
                    "command" = if ($pwshPath) { $pwshPath } else { "pwsh" }
                    "args" = @("-File", "`${workspaceFolder}/list_vms.ps1")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                },
                @{
                    "label" = "List Categories"
                    "type" = "shell"
                    "command" = if ($pwshPath) { $pwshPath } else { "pwsh" }
                    "args" = @("-File", "`${workspaceFolder}/list_categories.ps1")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                },
                @{
                    "label" = "Build Workbook"
                    "type" = "shell"
                    "command" = if ($pwshPath) { $pwshPath } else { "pwsh" }
                    "args" = @("-File", "`${workspaceFolder}/build_workbook.ps1")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                },
                @{
                    "label" = "Update VM Categories (PowerShell)"
                    "type" = "shell"
                    "command" = if ($pwshPath) { $pwshPath } else { "pwsh" }
                    "args" = @("-File", "`${workspaceFolder}/update_vm_categories.ps1")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                },
                @{
                    "label" = "Update VM Categories (Python)"
                    "type" = "shell"
                    "command" = if ($pythonPath) { $pythonPath } else { "python" }
                    "args" = @("update_categories_for_vm.py")
                    "group" = "build"
                    "presentation" = @{
                        "echo" = $true
                        "reveal" = "always"
                        "panel" = "new"
                    }
                    "options" = @{
                        "cwd" = "`${workspaceFolder}"
                    }
                }
            )
        }
        
        $tasksJson = $tasksConfig | ConvertTo-Json -Depth 5
        $tasksJson | Set-Content -Path $tasksPath -Encoding UTF8
        Write-Success "VS Code tasks configuration created at: $tasksPath"
        #>
        
        # Install recommended extensions
        Write-Info "Installing recommended VS Code extensions..."
        $extensions = @(
            "ms-python.python",
            "ms-vscode.powershell", 
            "redhat.vscode-yaml",
            "donjayamanne.githistory",
            "eamodio.gitlens"
        )
        
        foreach ($extension in $extensions) {
            try {
                Write-Info "Installing extension: $extension"
                $result = code --install-extension $extension --force 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Success "Extension installed: $extension"
                }
                else {
                    Write-Warning "Extension installation failed: $extension (Exit code: $LASTEXITCODE)"
                    Write-Warning "Output: $result"
                }
            }
            catch {
                Write-Warning "Could not install extension: $extension - $($_.Exception.Message)"
            }
        }
        Write-Success "VS Code extensions installation completed"
        
        # Open VS Code at repository root
        Write-Info "Opening VS Code at repository root..."
        Start-Process -FilePath "code" -ArgumentList "`"$RepositoryPath`"" -NoNewWindow
        Write-Success "VS Code opened at: $RepositoryPath"
        
    }
    catch {
        Write-Warning "Failed to configure VS Code: $($_.Exception.Message)"
        Write-Info "VS Code can still be used manually"
    }
}
else {
    if (-not $repositoryAvailable) {
        Write-Warning "Skipping VS Code configuration - repository not available"
    }
    if (-not (Test-CommandExists "code")) {
        Write-Warning "Skipping VS Code configuration - VS Code not available"
    }
}

#endregion

#region PowerShell Modules Installation

Write-Section "Installing PowerShell Modules"

try {
    # Install NuGet package provider first
    Write-Info "Installing NuGet package provider..."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
    Write-Success "NuGet package provider installed"
    
    # Check if ImportExcel module is installed
    $importExcelModule = Get-Module -ListAvailable -Name ImportExcel
    
    if (-not $importExcelModule -or $Force) {
        Write-Info "Installing ImportExcel module..."
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        Write-Success "ImportExcel module installed"
    }
    else {
        Write-Success "ImportExcel module is already installed (version: $($importExcelModule.Version))"
    }
    
    # Import the module to verify
    Import-Module ImportExcel -Force
    Write-Success "ImportExcel module imported successfully"
}
catch {
    Write-Error "Failed to install PowerShell modules: $($_.Exception.Message)"
}

#endregion

#region Python Virtual Environment Setup

Write-Section "Setting up Python Virtual Environment"

if (-not $repositoryAvailable) {
    Write-Warning "Skipping Python virtual environment setup - repository not available"
    Write-Info "Run the script again after cloning the repository"
}
else {
    try {
        # Change to repository directory
        Push-Location $RepositoryPath
    
    $venvPath = Join-Path $RepositoryPath ".venv"
    
    # Create virtual environment if it doesn't exist
    if (-not (Test-Path $venvPath) -or $Force) {
        Write-Info "Creating Python virtual environment..."
        python -m venv .venv
        Write-Success "Python virtual environment created"
    }
    else {
        Write-Success "Python virtual environment already exists"
    }
    
    # Activate virtual environment and install packages
    $activateScript = Join-Path $venvPath "Scripts\Activate.ps1"
    if (Test-Path $activateScript) {
        Write-Info "Activating virtual environment..."
        & $activateScript
        
        # Check if requirements.txt exists in files folder
        $requirementsPath = Join-Path $RepositoryPath "files\requirements.txt"
        
        if (Test-Path $requirementsPath) {
            Write-Info "Installing Python packages from requirements.txt..."
            python -m pip install --upgrade pip
            python -m pip install -r $requirementsPath
            Write-Success "Python packages installed from requirements.txt"
        }
        else {
            # Create requirements.txt with expected packages
            Write-Info "Creating requirements.txt with expected packages..."
            $requirementsDir = Join-Path $RepositoryPath "files"
            if (-not (Test-Path $requirementsDir)) {
                New-Item -Path $requirementsDir -ItemType Directory -Force | Out-Null
            }
            
            $PythonRequirements | Set-Content -Path $requirementsPath -Encoding UTF8
            Write-Info "Installing Python packages..."
            python -m pip install --upgrade pip
            python -m pip install -r $requirementsPath
            Write-Success "Python packages installed"
        }
        
        # Verify installations
        Write-Info "Verifying Python package installations..."
        $packages = @("requests", "pandas", "openpyxl", "urllib3")
        foreach ($package in $packages) {
            try {
                python -c "import $package; print('$package : ' + $package.__version__)"
            }
            catch {
                Write-Warning "Could not verify $package installation"
            }
        }
        
        Write-Success "Python virtual environment setup complete"
    }
    else {
        Write-Error "Could not find virtual environment activation script"
    }
    }
    catch {
        Write-Error "Failed to setup Python virtual environment: $($_.Exception.Message)"
    }
    finally {
        Pop-Location
    }
}

#endregion

#region Final Verification and Instructions

Write-Section "Installation Complete"

Write-Host ""
Write-Host " Environment setup completed successfully!" -ForegroundColor Green
Write-Host ""

Write-Host " Next Steps:" -ForegroundColor Cyan
Write-Host "1. Navigate to the repository directory:" -ForegroundColor White
Write-Host "   cd `"$RepositoryPath`"" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Edit the configuration file with your Nutanix details:" -ForegroundColor White
Write-Host "   notepad files\vars.txt" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Activate the Python virtual environment:" -ForegroundColor White
Write-Host "   .\.venv\Scripts\Activate.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "4. Run the scripts in order:" -ForegroundColor White
Write-Host "   .\list_vms.ps1" -ForegroundColor Gray
Write-Host "   .\list_categories.ps1" -ForegroundColor Gray
Write-Host "   .\build_workbook.ps1" -ForegroundColor Gray
Write-Host "   # Edit Excel file as needed" -ForegroundColor Gray
Write-Host "   .\update_vm_categories.ps1" -ForegroundColor Gray
Write-Host "   python update_categories_for_vm.py" -ForegroundColor Gray
Write-Host ""

Write-Host " Documentation:" -ForegroundColor Cyan
Write-Host "   Repository: https://github.com/hardevsanghera/ntnx-v4api-cats" -ForegroundColor Gray
Write-Host "   README.md contains detailed usage instructions" -ForegroundColor Gray
Write-Host ""

Write-Host "  Important Notes:" -ForegroundColor Yellow
Write-Host "    Update files\vars.txt with your Nutanix Prism Central details" -ForegroundColor Gray
Write-Host "    Scripts use plain-text passwords - secure appropriately" -ForegroundColor Gray
Write-Host "    SSL certificate checking is disabled - modify for production" -ForegroundColor Gray
Write-Host "    Microsoft Excel is required for COM automation features" -ForegroundColor Gray
Write-Host ""

# Create a summary report
$summaryReport = @"
=== Nutanix v4 API Environment Setup Summary ===
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Repository: $RepositoryPath

Installed Components:
- PowerShell 7: $(if ($needsPwsh) { "Installed" } else { "Already present" })
- Python 3.13+: $(if ($needsPython) { "Installed" } else { "Already present" })
- Visual Studio Code: $(if (Test-CommandExists "code") { "Available" } else { "Installation attempted" })
- Git for Windows: $(if (Test-CommandExists "git") { "Available" } else { "Installation attempted" })
- ImportExcel Module: Installed/Verified
- Python Virtual Environment: Created at $RepositoryPath\.venv
- Python Packages: Installed from requirements

Repository Status: $(if (Test-Path $RepositoryPath) { "Ready" } else { "Setup required" })

Next: Configure files\vars.txt and follow the README.md workflow
"@

$summaryPath = Join-Path $env:TEMP "setup-summary.txt"
try {
    $summaryReport | Set-Content -Path $summaryPath -Encoding UTF8
    Write-Host " Setup summary saved to: $summaryPath" -ForegroundColor Green
}
catch {
    Write-Warning "Could not save setup summary file"
}

Write-Host ""
Write-Success "Setup completed! Ready to work with Nutanix v4 APIs."

#endregion
 
