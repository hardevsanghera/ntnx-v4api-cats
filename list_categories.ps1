<#
Simple script to fetch categories from Prism Central and save JSON to .\scratch\categories.json
Created as a copy/simplified variant of new3gemenips.ps1 per user request.
#>

# --- API Configuration (loaded from vars.txt) ---
# Read variables from vars.txt file in the files folder
$varsFile = Join-Path -Path $PWD -ChildPath 'files\vars.txt'
if (-not (Test-Path -Path $varsFile)) {
    Write-Error "vars.txt file not found at $varsFile"
    exit 1
}

$vars = @{}
Get-Content $varsFile | ForEach-Object {
    if ($_ -match '^([^=]+)=(.*)$') {
        $vars[$matches[1]] = $matches[2]
    }
}

$baseUrl = $vars['baseUrl']
$username = $vars['username']
$password = $vars['password']

$uri = "$baseUrl/prism/v4.1/config/categories?$limit=125"

# Build Basic Auth header
$base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$($password)"))
$headers = @{ Authorization = "Basic $base64AuthInfo" }

# Target output path
$scratchDir = Join-Path -Path $PWD -ChildPath 'scratch'
$outFile = Join-Path -Path $scratchDir -ChildPath 'categories.json'

try {
    Write-Host "Calling categories endpoint: $uri" -ForegroundColor Cyan

    # Perform GET request. Keep -SkipCertificateCheck and -SkipHttpErrorCheck for parity with original scripts
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

    if ($response) {
        if (-not (Test-Path -Path $scratchDir)) { New-Item -Path $scratchDir -ItemType Directory -Force | Out-Null }

        try {
            $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $outFile -Encoding utf8 -Force
            Write-Host "Saved categories JSON to $outFile" -ForegroundColor Yellow
        } catch {
            Write-Error "Failed to write $outFile : $($_.Exception.Message)"
            exit 1
        }
    } else {
        Write-Warning "No response received from $uri"
        exit 1
    }
} catch {
    Write-Error "An error occurred while calling the categories endpoint: $($_.Exception.Message)"
    exit 1
}
