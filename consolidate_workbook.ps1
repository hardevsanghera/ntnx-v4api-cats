<#
.SYNOPSIS
    Script Name: consolidate_workbook.ps1
    Author: hardev@nutanix.com + Co-Pilot
    Date: October 2025
    Version: 1.0

.DESCRIPTION
    Simple script to fetch categories from Prism Central and save JSON to .\scratch\categories.json
    PowerShell 7 script that:
        1) Opens $PWD\files\VMsToUpdate_SKEL.xlsx
        2) Appends all sheets from $PWD\scratch\cat_map.xlsx into the opened workbook
        3) Saves the updated workbook to $PWD\scratch\VMsToUpdate.xlsx (don't use the timestamp switch)
        4) Closes all files and quits Excel

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

try {
    $root = $PWD
    $skelPath = Join-Path $root 'files\VMsToUpdate_SKEL.xlsx'
    $mapPath = Join-Path $root 'scratch\cat_map.xlsx'
    $outDir = Join-Path $root 'scratch'

    if (-not (Test-Path $skelPath)) { throw "SKEL workbook not found: $skelPath" }
    if (-not (Test-Path $mapPath)) { throw "Mapping workbook not found: $mapPath" }
    if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory -Force | Out-Null }

    $ts = (Get-Date).ToString('yyyy-MM-dd-HH-mm')
    #$outFile = "VMsToUpdate-$ts.xlsx"
    $outFile = "VMsToUpdate-PROD.xlsx"
    $outPath = Join-Path $outDir $outFile

    Write-Host "Starting workbook consolidation..."
    Write-Host "Opening Excel (COM)..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Opening destination workbook: $skelPath"
    $wbDest = $excel.Workbooks.Open($skelPath)

    Write-Host "Opening source workbook to add: $mapPath"
    $wbSrc = $excel.Workbooks.Open($mapPath)

    $count = $wbSrc.Worksheets.Count
    Write-Host "Copying $count worksheet(s) from cat_map into destination workbook..."

    $missing = [System.Type]::Missing
    for ($i = 1; $i -le $count; $i++) {
        $ws = $wbSrc.Worksheets.Item($i)
        try {
            $lastIndex = $wbDest.Worksheets.Count
            if ($lastIndex -ge 1) {
                $after = $wbDest.Worksheets.Item($lastIndex)
                try {
                    # Try native copy first
                    $ws.Copy($missing, $after)
                } catch {
                    Write-Warning "Worksheet.Copy failed for '$($ws.Name)' (index $i). Falling back to range copy: $($_.Exception.Message)"
                    try {
                        $newSheet = $wbDest.Worksheets.Add($missing, $after)
                        try { $newSheet.Name = $ws.Name } catch { $newSheet.Name = "$($ws.Name)-copy-$i" }
                        $srcRange = $ws.UsedRange
                        if ($srcRange -and $srcRange.Count -gt 0) {
                            $srcRange.Copy()
                            $destRange = $newSheet.Range('A1')
                            $newSheet.Paste($destRange)
                            $excel.CutCopyMode = $false
                        }
                    } finally {
                        if ($srcRange) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($srcRange) | Out-Null }
                        if ($destRange) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($destRange) | Out-Null }
                        if ($newSheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($newSheet) | Out-Null }
                        $srcRange = $null; $destRange = $null; $newSheet = $null
                    }
                } finally {
                    if ($after) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($after) | Out-Null }
                    $after = $null
                }
            } else {
                # No sheets in destination (unlikely) - try simple copy then fallback
                try { $ws.Copy() } catch {
                    Write-Warning "Worksheet.Copy (no-destination) failed for '$($ws.Name)' (index $i). Falling back to range copy: $($_.Exception.Message)"
                    try {
                        $newSheet = $wbDest.Worksheets.Add()
                        try { $newSheet.Name = $ws.Name } catch { $newSheet.Name = "$($ws.Name)-copy-$i" }
                        $srcRange = $ws.UsedRange
                        if ($srcRange -and $srcRange.Count -gt 0) {
                            $srcRange.Copy()
                            $destRange = $newSheet.Range('A1')
                            $newSheet.Paste($destRange)
                            $excel.CutCopyMode = $false
                        }
                    } finally {
                        if ($srcRange) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($srcRange) | Out-Null }
                        if ($destRange) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($destRange) | Out-Null }
                        if ($newSheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($newSheet) | Out-Null }
                        $srcRange = $null; $destRange = $null; $newSheet = $null
                    }
                }
            }
        } catch {
            Write-Warning "Failed to copy worksheet '$($ws.Name)' (index $i): $($_.Exception.Message)"
        } finally {
            if ($ws) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null }
            $ws = $null
        }
    }

    # Make 'ToUpdate' worksheet the active sheet if it exists
    try {
        $toUpdateSheet = $null
        try { $toUpdateSheet = $wbDest.Worksheets.Item('ToUpdate') } catch { $toUpdateSheet = $null }
        if ($toUpdateSheet) {
            try {
                $toUpdateSheet.Activate()
                # Optionally select A1 to ensure a visible active cell
                try { $toUpdateSheet.Range('A1').Select() } catch { }
                Write-Host "Activated worksheet 'ToUpdate' as the active sheet." -ForegroundColor Green
            } catch {
                Write-Warning "Failed to activate 'ToUpdate' sheet: $($_.Exception.Message)"
            } finally {
                if ($toUpdateSheet) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($toUpdateSheet) | Out-Null }
                $toUpdateSheet = $null
            }
        } else {
            Write-Host "Worksheet 'ToUpdate' not found; leaving active sheet unchanged." -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Error while trying to activate 'ToUpdate' sheet: $($_.Exception.Message)"
    }

    Write-Host "Saving updated workbook to: $outPath"
    $wbDest.SaveAs($outPath, 51)

    Write-Host "Closing workbooks..."
    if ($wbSrc) { $wbSrc.Close($false) }
    if ($wbDest) { $wbDest.Close($false) }
    if ($excel) { $excel.Quit() }

    # Release COM objects
    if ($wbSrc) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbSrc) | Out-Null }
    if ($wbDest) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wbDest) | Out-Null }
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    Write-Host "Done. Saved: $outPath"
    exit 0

} catch {
    Write-Error "Error: $_"
    try {
        if ($wbSrc) { $wbSrc.Close($false) }
        if ($wbDest) { $wbDest.Close($false) }
        if ($excel) { $excel.Quit() }
    } catch { }
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    exit 1
}
