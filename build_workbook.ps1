<#
.SYNOPSIS
    Script Name: build_workbook.ps1
    Author: hardev@nutanix.com + Co-Pilot
    Date: October 2025
    Version: 1.0

.DESCRIPTION
    This script processes Nutanix v4 API VM and category data to create comprehensive reports
    It reads JSON files containing VM configurations and category definitions from the scratch directory
    Outputs include CSV files and Excel workbooks with VM-to-category mappings for analysis and management

    Requires: Windows with Excel installed (uses COM interop). Run with PowerShell 7 (pwsh).
    
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

function Resolve-CategoryMappings {
    param(
        [string]$VmJsonPath =       "$PWD\scratch\vm_list.json",
        [string]$CategoryJsonPath = "$PWD\scratch\categories.json",
        [string]$OutCsvPath =       "$PWD\scratch\vm_categories.csv",
        [switch]$SplitCategories = $true,                # if set, write separate columns per category key DO NOT CHANGE THIS from $true
        [switch]$TimestampExcel = $false                 # if set, append timestamp to Excel filename
    )

    if (-not (Test-Path $VmJsonPath)) {
        Write-Warning "VM JSON file not found at $VmJsonPath"
        return
    }

    $data = (Get-Content -Path $VmJsonPath -Raw | ConvertFrom-Json)

    # Attempt to load category definitions from provided path; fall back to objects in the VM dump
    $catDefs = @()
    if (Test-Path $CategoryJsonPath) {
        try { $catDefs = (Get-Content -Path $CategoryJsonPath -Raw | ConvertFrom-Json).data } catch { $catDefs = @() }
    }
    if ( ( -not $catDefs -or $catDefs.Count -lt 1 ) -and $data.data) {
        $catDefs = $data.data | Where-Object { $_.'$objectType' -eq 'prism.v4.config.Category' }
    }

    # Build a mapping of category extId -> category object for quick lookup
    $catByExt = @{}
    if ($catDefs) { foreach ($c in $catDefs) { if ($c.extId) { $catByExt[$c.extId] = $c } } }

    $results = @()
    foreach ($vm in ($data.data | Where-Object { $_.'$objectType' -eq 'vmm.v4.ahv.config.Vm' })) {
        $vmName = if ($vm.status -and $vm.status.name) { $vm.status.name } elseif ($vm.name) { $vm.name } else { '<unnamed>' }
        $vmExt = if ($vm.metadata -and $vm.metadata.extId) { $vm.metadata.extId } elseif ($vm.extId) { $vm.extId } else { '<no-extId>' }

        $resolved = @()
        $catMap = @{}
        if ($vm.categories) {
            foreach ($cref in $vm.categories) {
                $cext = $cref.extId
                if ($cext) {
                    $cobj = $null
                    if ($catByExt.ContainsKey($cext)) { $cobj = $catByExt[$cext] }
                    if (-not $cobj) { $cobj = $catDefs | Where-Object { $_.extId -eq $cext } | Select-Object -First 1 }
                    if ($cobj) {
                        $resolved += "{0}={1}" -f $cobj.key, $cobj.value
                        $k = $cobj.key
                        if (-not $catMap.ContainsKey($k)) { $catMap[$k] = @() }
                        $catMap[$k] += $cobj.value
                    } else {
                        $resolved += $cext
                        # Use the extId itself as a fallback key
                        $k = $cext
                        if (-not $catMap.ContainsKey($k)) { $catMap[$k] = @() }
                        $catMap[$k] += $cext
                    }
                }
            }
        }

        $results += [PSCustomObject]@{
            'VM Name' = $vmName
            'VM extId' = $vmExt
            Categories = ($resolved -join '; ')
            CategoriesMap = $catMap
        }
    }

    # Print table and save CSV
    if ($results.Count -eq 0) { 
        Write-Host "No VM category mappings found." 
    } else { 
        $results | Format-Table -AutoSize 
    }

    # Ensure output directories exist
    $csvDir = Split-Path -Path $OutCsvPath -Parent
    if (-not (Test-Path $csvDir)) { New-Item -Path $csvDir -ItemType Directory -Force | Out-Null }

    try {
        if ($SplitCategories) {
            # Build column list of unique category keys (keep order stable)
            $allKeys = @()
            $allKeys += ($catDefs | Where-Object { $_.key } | Select-Object -ExpandProperty key)
            $allKeys += ($results | ForEach-Object { if ($_.CategoriesMap) { $_.CategoriesMap.Keys } })
            $allKeys = $allKeys | Select-Object -Unique

            # Build rows where each category key becomes a column
            # IMPORTANT: Original result objects store properties literally as 'VM Name' and 'VM extId'.
            # Previous code incorrectly referenced $r.Name / $r.ExtId which are NULL, yielding blank columns.
            $csvRows = foreach ($r in $results) {
                $h = @{ 'VM Name' = $r.'VM Name'; 'VM extId' = $r.'VM extId' }
                foreach ($k in $allKeys) { $h[$k] = '' }
                if ($r.CategoriesMap) {
                    foreach ($k in $allKeys) {
                        if ($r.CategoriesMap.ContainsKey($k)) { $h[$k] = ($r.CategoriesMap[$k] -join '; ') }
                    }
                }
                [PSCustomObject]$h
            }

            $propNames = @('VM Name','VM extId') + $allKeys
            $csvRows | Select-Object $propNames | Export-Csv -Path $OutCsvPath -NoTypeInformation -Encoding UTF8
    
        } else {
            # Legacy CSV format: Name, Categories, VM extId
            $csvRows = $results | ForEach-Object {
                [PSCustomObject]@{
                    'VM Name' = $_.'VM Name'
                    'VM extId' = $_.'VM extId'
                    Categories = $_.Categories
                }
            }
            # Desired order: Name, VM extId, Categories
            $csvRows | Select-Object 'VM Name','VM extId',Categories | Export-Csv -Path $OutCsvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Saved VM categories to $OutCsvPath" -ForegroundColor Yellow
    } catch {
        Write-Warning "Failed to save CSV: $($_.Exception.Message)"
    }

    # Also write to an Excel workbook at $HOME\Documents\V4APIs\cat_map_{timestamp}.xlsx (if requested)
    $ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
    if ($TimestampExcel) {
        $excelFilename = "cat_map_$ts.xlsx"
    } else {
        $excelFilename = "cat_map.xlsx"
    }
    $excelPath = Join-Path $PWD (Join-Path '\scratch\' $excelFilename)
    $excelDir = Split-Path -Path $excelPath -Parent
    if (-not (Test-Path $excelDir)) { New-Item -Path $excelDir -ItemType Directory | Out-Null }

    # Start fresh each run: remove any existing workbook to avoid leftover/legacy sheets
    if (Test-Path $excelPath) {
        try {
            Remove-Item -Path $excelPath -Force -ErrorAction Stop
        } catch {
            Write-Warning "Failed to remove existing workbook at $excelPath : $($_.Exception.Message)"
        }
    }

    # Prefer ImportExcel's Export-Excel if available
    if (Get-Command -Name Export-Excel -ErrorAction SilentlyContinue) {
        try {
            if ($SplitCategories) {
                # Build column list of unique category keys (keep order stable)
                $allKeys = @()
                $allKeys += ($catDefs | Where-Object { $_.key } | Select-Object -ExpandProperty key)
                $allKeys += ($results | ForEach-Object { if ($_.CategoriesMap) { $_.CategoriesMap.Keys } })
                $allKeys = $allKeys | Select-Object -Unique

                # Create ordered objects where each key becomes a column, using per-VM CategoriesMap
                $splitRows = foreach ($r in $results) {
                    $h = @{ 'VM Name' = $r.'VM Name'; 'VM extId' = $r.'VM extId' }
                    foreach ($k in $allKeys) { $h[$k] = '' }
                    if ($r.CategoriesMap) {
                        foreach ($k in $allKeys) {
                            if ($r.CategoriesMap.ContainsKey($k)) { $h[$k] = ($r.CategoriesMap[$k] -join '; ') }
                        }
                    }
                    [PSCustomObject]$h
                }
                # New order: Name, VM extId, <category keys>
                $propNames = @('VM Name','VM extId') + $allKeys
                $ordered = $splitRows | Select-Object $propNames
            } else {
                # Reorder properties to Name, VM extId, Categories for output (legacy non-split)
                $ordered = $results | Select-Object @{n='VM Name';e={$_."VM Name"}}, @{n='VM extId';e={$_."VM extId"}}, @{n='Categories';e={$_.Categories}}
            }
            # -AutoSize and -AutoFilter improve readability; headers will be bold by default in Export-Excel table
            $null = $ordered | Export-Excel -Path $excelPath -WorksheetName 'VMCategories' -AutoSize -AutoFilter -TableName 'VMCategories'
    
            # Also write AllCategories sheet: two columns 'Category' and 'extID'
            #try {
                $allCatRows = @()
                if ($catDefs) {
                    foreach ($c in $catDefs) {
                        $catKey   = if ($c.key) { $c.key } else { $null }
                        $catValue = if ($c.value) { $c.value } else { '' }
                        $ext      = if ($c.extId) { $c.extId } else { '' }
                        # Category column = key if present else value (legacy behavior); Value column always = value
                        $categoryCol = if ($catKey) { $catKey } else { $catValue }
                        $allCatRows += [PSCustomObject]@{ Category = $categoryCol; Value = $catValue; extID = $ext }
                    }
                } elseif ($catByExt.Keys) {
                    foreach ($k in $catByExt.Keys) {
                        $c = $catByExt[$k]
                        $catKey   = if ($c.key) { $c.key } else { $null }
                        $catValue = if ($c.value) { $c.value } else { '' }
                        $categoryCol = if ($catKey) { $catKey } else { $catValue }
                        $allCatRows += [PSCustomObject]@{ Category = $categoryCol; Value = $catValue; extID = $k }
                    }
                }

                if ($allCatRows.Count -gt 0) {
                    # If any row is missing Value (older objects), normalize them
                    $allCatRows = $allCatRows | ForEach-Object {
                        $val = if ($_.PSObject.Properties.Match('Value').Count -gt 0) { $_.Value } else { '' }
                        [PSCustomObject]@{ Category = $_.Category; Value = $val; extID = $_.extID }
                    }
                    # Export-Excel will append a sheet with columns Category, Value, extID
                    $null = $allCatRows | Export-Excel -Path $excelPath -WorksheetName 'AllCategories' -AutoSize -AutoFilter -TableName 'AllCategories' -Append
               }  
            
                #} catch {
                #    Write-Warning "Failed to write AllCategories sheet via Export-Excel: $($_.Exception.Message)"
                #}

            Write-Host "Saved Excel workbook to $excelPath" -ForegroundColor Yellow
        } catch {
            Write-Warning "Export-Excel failed: $($_.Exception.Message)"
        }
   } 
  # catch {
   <#
   .SYNOPSIS
   Short description
   
   .DESCRIPTION
   Long description
   
   .PARAMETER VmJsonPath
   Parameter description
   
   .PARAMETER CategoryJsonPath
   Parameter description
   
   .PARAMETER OutCsvPath
   Parameter description
   
   .PARAMETER SplitCategories
   Parameter description
   
   .PARAMETER TimestampExcel
   Parameter description
   
   .EXAMPLE
   An example
   
   .NOTES
   General notes
   #>    <#
   .SYNOPSIS
   Short description
   
   .DESCRIPTION
   Long description
   
   .PARAMETER VmJsonPath
   Parameter description
   
   .PARAMETER CategoryJsonPath
   Parameter description
   
   .PARAMETER OutCsvPath
   Parameter description
   
   .PARAMETER SplitCategories
   Parameter description
   
   .PARAMETER TimestampExcel
   Parameter description
   
   .EXAMPLE
   An example
   
   .NOTES
   General notes
   #>#Write-Warning "An error occurred while attempting to use Export-Excel: $($_.Exception.Message)"
   }
#} # Closing the if block for Export-Excel availability
#else {
#       Write-Warning "Export-Excel is not available, and no fallback logic is implemented."
#   }
#        # Fallback: no Export-Excel available, delegate COM workbook creation to helper script
#        try {
#            $scriptPath = Join-Path $PSScriptRoot 'scripts\create-excel-workbook.ps1'
#            if (-not (Test-Path $scriptPath)) { throw "Helper script not found: $scriptPath" }
#
#            # Choose which data collection to pass: if we built an ordered object for Export-Excel use that, otherwise pass results
#            $toPass = if ($ordered) { $ordered } else { $results }
#
#            & pwsh -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Path $excelPath -Results $toPass -CatDefs $catDefs -CatByExt $catByExt -SplitCategories:($SplitCategories.IsPresent)
#        } catch {
#            Write-Warning "Failed to write Excel workbook via helper script: $($_.Exception.Message)"
#        }
#    }
#}

# Entry point: call Resolve-CategoryMappings after the script's usual save point
try {
    Resolve-CategoryMappings
} catch {
    Write-Warning "An error occurred in Resolve-CategoryMappings: $($_.Exception.Message)"
}

 # Closing the Resolve-CategoryMappings function block

# Prints whether Export-Excel is available and module/version if present
if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
  Write-Host 'Export-Excel: available'
  Get-Command Export-Excel | Select-Object Name, Source, ModuleName, Version | Format-List
} else {
  Write-Host 'Export-Excel: missing'
  Get-Module -ListAvailable ImportExcel | ForEach-Object { Write-Host ("ImportExcel module found: {0} {1}" -f $_.Name, $_.Version) }
}