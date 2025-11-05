<#
.SYNOPSIS
    Script Name: list_vms.ps1
    Author: hardev@nutanix.com + Co-Pilot
    Date: October 2025
    Version: 1.0

.DESCRIPTION
    This script retrieves a list of VMs from a Nutanix cluster using the REST API,
    saves the response to a JSON file, and creates a mapping of VM names to their extIds.
    The mapping is printed to the console and also saved to a CSV file.
    
    NB:
    This script is provided "AS IS" without warranty of any kind.
    Use of this script is at your own risk. 
    The author(s) make no representations or warranties, express or implied, 
    regarding the script’s functionality, fitness for a particular purpose, 
    or reliability. 

    By using this script, you agree that you are solely responsible 
    for any outcomes, including loss of data, system issues, or 
    other damages that may result from its execution. 
    No support or maintenance is provided.

.NOTES
    You may copy, edit, customize and use as needed. Test thoroughly in a safe environment 
    before deploying to production systems.
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

# --- Create the Basic Authentication Header ---
$base64AuthInfo = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes("$($username):$($password)"))
$headers = @{
    Authorization = "Basic $base64AuthInfo"
}

# --- Make the REST API Call with Pagination ---
try {
    $pageSize = 100  # Maximum allowed by the API
    $pageNumber = 0
    $allVMs = @()
    $hasMorePages = $true

    while ($hasMorePages) {
        $uri = "$baseUrl/vmm/v4.1/ahv/config/vms?`$limit=$pageSize&`$page=$pageNumber"
        Write-Host "Calling API at URI: $uri (Page $pageNumber)" -ForegroundColor Cyan

        $response = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers -SkipCertificateCheck -SkipHttpErrorCheck

        if ($response -and $response.data) {
            $currentPageVMs = $response.data
            $vmCount = ($currentPageVMs | Measure-Object).Count
            Write-Host "Retrieved $vmCount VMs from page $pageNumber" -ForegroundColor Green
            
            $allVMs += $currentPageVMs
            
            # Check if there are more pages
            if ($vmCount -lt $pageSize) {
                $hasMorePages = $false
                Write-Host "Reached last page. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            } else {
                $pageNumber++
            }
        } else {
            Write-Host "No more data returned. Total VMs retrieved: $($allVMs.Count)" -ForegroundColor Yellow
            $hasMorePages = $false
        }
    }

    if ($allVMs.Count -gt 0) {
        Write-Host "Successfully retrieved all $($allVMs.Count) VMs." -ForegroundColor Green

        $vmFile = Join-Path -Path $PWD -ChildPath '\scratch\vm_list.json'

        # Create a combined response object to save
        $combinedResponse = @{
            data = $allVMs
            totalCount = $allVMs.Count
        }

        # Save the combined API response to vm_list.json (overwrite)
        try {
            $combinedResponse | ConvertTo-Json -Depth 10 | Out-File -FilePath $vmFile -Encoding utf8 -Force
            Write-Host "Saved all VMs to $vmFile" -ForegroundColor Yellow
        } catch {
            Write-Error "Failed to write $vmFile :"
            exit 1
        }

        # Read the JSON file (expected structure with entities array)
        try {
            $jsonText = Get-Content -Path $vmFile -Raw -ErrorAction Stop

            # Quick sanity check: JSON must start with '{' or '[' after optional BOM and whitespace
            $snippet = $jsonText.TrimStart([char]0xFEFF).Substring(0, [Math]::Min(1000, $jsonText.Length)) -replace "\r?\n", " `n "  
            $firstNonWs = ($jsonText.TrimStart([char]0xFEFF) -match "^\s*(.)" ) | Out-Null; $firstChar = $Matches[1]
            if (-not ($firstChar -eq '{' -or $firstChar -eq '[')) {
                Write-Error "The file '$vmFile' does not appear to contain JSON. First non-whitespace character: '$firstChar'"
                Write-Host "--- File snippet (first 1k chars) ---" -ForegroundColor Yellow
                Write-Host $snippet
                # Save snippet for debugging
                $dbgFile = Join-Path -Path $PWD -ChildPath 'scratch\vm_list_debug_snippet.txt'
                try {
                    $snippet | Out-File -FilePath $dbgFile -Encoding utf8 -Force
                    Write-Host "Saved debug snippet to $dbgFile" -ForegroundColor Yellow
                } catch {
                    Write-Warning "Failed to write debug snippet: $($_.Exception.Message)"
                }
                Write-Error "Aborting: vm_list.json is not valid JSON. Inspect the debug snippet file or the saved API response."
                exit 1
            }

            $data = $jsonText | ConvertFrom-Json -ErrorAction Stop
        } catch {
            Write-Error "Failed to read or parse $vmFile :"
            Write-Error ($_.Exception.ToString())
            exit 1
        }

    # Build a hashtable mapping Name -> array of extIds (or uuid fallback)
    $vmHash = @{}

        if ($null -ne $data.data) {
            $items = $data.data
        } else {
            # maybe the JSON is the array itself
            $items = $data
        }

        foreach ($item in $items) {
            # Determine name (common locations)
            $name = $null
            if ($null -ne $item.status -and $null -ne $item.status.name) { $name = $item.status.name }
            elseif ($null -ne $item.name) { $name = $item.name }

            # Determine extId (many Nutanix objects use metadata.extId or metadata.uuid)
            $extId = $null
            if ($null -ne $item.metadata) {
                if ($null -ne $item.metadata.extId) { $extId = $item.metadata.extId }
                elseif ($null -ne $item.metadata.uuid) { $extId = $item.metadata.uuid }
            } elseif ($null -ne $item.extId) { $extId = $item.extId }

            if ($name) {
                if (-not $vmHash.ContainsKey($name)) {
                    $vmHash[$name] = @()
                }
                $vmHash[$name] += $extId
            }
        }

        # Print the hashtable entries
        Write-Host "VM Name -> extId mapping:" -ForegroundColor Cyan
        if ($vmHash.Keys.Count -eq 0) {
            Write-Host "No VM entries found in $vmFile" -ForegroundColor Yellow
        } else {
            $vmHash.GetEnumerator() | Sort-Object 'VM Name' | ForEach-Object {
                $k = $_.Key
                $v = $_.Value
                $joined = ""
                # FIX 1: Correcting the if statement for array joining
                if ($v -is [System.Array]) {
                    $joined = ($v | Where-Object { $_ -ne $null }) -join '; '
                } else {
                    $joined = ($v -ne $null ? $v : '<null>')
                }
                Write-Host ("{0} -> {1}" -f $k, ($joined -ne '' ? $joined : '<null>'))
            }
            Write-Output "*** vmhash"
            Write-Output $vmHash

            # Also export mapping to CSV
            $csvFile = Join-Path -Path $PWD -ChildPath 'scratch\vm_mapping.csv'
            $csvRecords = $vmHash.GetEnumerator() | ForEach-Object {
                $extIdsValue = ""
                # FIX 2: Correcting the if statement inside the PSCustomObject
                if ($_.Value -is [System.Array]) {
                    $extIdsValue = ($_.Value | Where-Object { $_ -ne $null }) -join '; '
                } else {
                    $extIdsValue = $_.Value
                }
                [PSCustomObject]@{
                    'VM Name' = $_.Key
                    ExtIds = $extIdsValue
                }
            }
            Write-Output "*** csvrecords"
            Write-Output $csvRecords
            try {
                $csvRecords | Export-Csv -Path $csvFile -NoTypeInformation -Encoding utf8 -Force
                Write-Host "Saved mapping to $csvFile" -ForegroundColor Yellow
            } catch {
                Write-Warning "Failed to save mapping to CSV: $($_.Exception.Message)"
            }
        }

    } else {
        Write-Error "API call failed. No response received."
    }

} catch {
    Write-Error "An error occurred during the API call."
    # Print full exception details to help debugging
    Write-Error ($_.Exception.ToString())
}
