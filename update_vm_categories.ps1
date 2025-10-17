#!/usr/bin/env pwsh
<#

Requires: ImportExcel module.
#>
[CmdletBinding()]
param(
    [string]$Workbook = 'scratch/VMsToUpdate-PROD.xlsx',
    [switch]$DryRun,
    [switch]$IgnoreCase,
    [int]$Verbosity = 0
)

function Write-Verbose2 { param([string]$Message,[int]$Level=1) if ($Verbosity -ge $Level) { Write-Host "[v$Level] $Message" -ForegroundColor DarkCyan } }

# --- Module check ---
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Error "ImportExcel module not installed. Install with: Install-Module ImportExcel -Scope CurrentUser"; exit 2 }
Import-Module ImportExcel -ErrorAction Stop

# --- Workbook existence ---
if (-not (Test-Path -LiteralPath $Workbook)) { Write-Error "Workbook not found: $Workbook"; exit 2 }

# --- Helper: normalize (string) ---
function Normalize {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    $s = [string]$Value
    if ($IgnoreCase) { return $s.ToLowerInvariant() } else { return $s }
}

# --- Load data tables ---
# ToUpdate can be read via Import-Excel (no duplicate headers by case expected)
try { $toUpdateData = Import-Excel -Path $Workbook -Worksheet 'ToUpdate' }
catch { Write-Error "Failed to read 'ToUpdate' sheet: $($_.Exception.Message)"; exit 2 }

# VMCategories and AllCategories are read by cell to preserve header case and avoid collisions
function Read-WorksheetRows {
    param([string]$Path,[string]$Worksheet)
    try { $pkg = Open-ExcelPackage -Path $Path -ErrorAction Stop } catch { throw }
    try {
        $ws = $pkg.Workbook.Worksheets[$Worksheet]
        if (-not $ws) { throw "Worksheet not found: $Worksheet" }
        if (-not $ws.Dimension) { return @{ Headers=@(); Rows=@() } }
        $endRow = $ws.Dimension.End.Row
        $endCol = $ws.Dimension.End.Column
        $headers = for ($c=1;$c -le $endCol;$c++){ [string]$ws.Cells[1,$c].Text }
        $rows = @()
        for ($r=2;$r -le $endRow;$r++) {
            $vals = New-Object string[] $endCol
            for ($c=1;$c -le $endCol;$c++){ $vals[$c-1] = [string]$ws.Cells[$r,$c].Text }
            $rows += ,$vals
        }
        return @{ Headers=$headers; Rows=$rows }
    } finally { try { if ($pkg) { Close-ExcelPackage $pkg } } catch { } }
}

try {
    $vmSheet  = Read-WorksheetRows -Path $Workbook -Worksheet 'VMCategories'
    $allSheet = Read-WorksheetRows -Path $Workbook -Worksheet 'AllCategories'
} catch { Write-Error "Failed to read required sheets: $($_.Exception.Message)"; exit 2 }

# Build objects from VMCategories (only VM Name and VM extId)
$idxVmName = [Array]::IndexOf($vmSheet.Headers,'VM Name')
$idxVmExt  = [Array]::IndexOf($vmSheet.Headers,'VM extId')
if ($idxVmName -lt 0 -or $idxVmExt -lt 0) { Write-Error "Missing 'VM Name' or 'VM extId' in VMCategories"; exit 2 }
$vmCatData = foreach ($row in $vmSheet.Rows) {
    [PSCustomObject]@{ 'VM Name' = $row[$idxVmName]; 'VM extId' = $row[$idxVmExt] }
}

# Build objects from AllCategories (Category, Value, optional extID/extId)
$idxCat    = [Array]::IndexOf($allSheet.Headers,'Category')
$idxVal    = [Array]::IndexOf($allSheet.Headers,'Value')
$idxExtID  = [Array]::IndexOf($allSheet.Headers,'extID')
if ($idxExtID -lt 0) { $idxExtID = [Array]::IndexOf($allSheet.Headers,'extId') }
if ($idxCat -lt 0 -or $idxVal -lt 0) { Write-Error "Missing 'Category' or 'Value' in AllCategories"; exit 2 }
$allCatData = foreach ($row in $allSheet.Rows) {
    $o = [ordered]@{ Category = $row[$idxCat]; Value = $row[$idxVal] }
    if ($idxExtID -ge 0) { $o['extID'] = $row[$idxExtID] }
    [PSCustomObject]$o
}

$requiredColsToUpdate = 'VM Name','VM extId','UPDATE WITH CATEGORIES'
foreach ($c in $requiredColsToUpdate) { if (-not ($toUpdateData | Get-Member -Name $c -MemberType NoteProperty)) { Write-Error "Missing column '$c' in ToUpdate"; exit 2 } }
foreach ($c in 'VM Name','VM extId') { if (-not ($vmCatData | Get-Member -Name $c -MemberType NoteProperty)) { Write-Error "Missing column '$c' in VMCategories"; exit 2 } }
foreach ($c in 'Category','Value') { if (-not ($allCatData | Get-Member -Name $c -MemberType NoteProperty)) { Write-Error "Missing column '$c' in AllCategories"; exit 2 } }

# Detect extID column in AllCategories
$catExtIdHeader = @('extID','extId') | Where-Object { $allCatData | Get-Member -Name $_ -MemberType NoteProperty } | Select-Object -First 1
if (-not $catExtIdHeader) { Write-Warning "No extID/extId column found in AllCategories. Category UUID(s) values will be blank." }

# Build VM pair counts
$vmPairCounts = @{}
$vmFirstRow   = @{}
for ($i=0; $i -lt $vmCatData.Count; $i++) {
    $row = $vmCatData[$i]
    $vmName  = $row.'VM Name'
    $vmExtId = $row.'VM extId'
    if ([string]::IsNullOrWhiteSpace($vmName) -and [string]::IsNullOrWhiteSpace($vmExtId)) { continue }
    $key = (Normalize $vmName) + '||' + (Normalize $vmExtId)
    $vmPairCounts[$key] = 1 + ($vmPairCounts[$key] | ForEach-Object { $_ })
    if (-not $vmFirstRow.ContainsKey($key)) { $vmFirstRow[$key] = $i }
}

# Build Category/Value -> index and extID mapping (first occurrence rule)
$catRowIndex = @{}
$catUuidMap  = @{}
for ($i=0; $i -lt $allCatData.Count; $i++) {
    $row = $allCatData[$i]
    $cat = $row.Category; $val = $row.Value
    if ([string]::IsNullOrWhiteSpace($cat) -and [string]::IsNullOrWhiteSpace($val)) { continue }
    $key = (Normalize $cat)+'||'+(Normalize $val)
    if (-not $catRowIndex.ContainsKey($key)) {
        $catRowIndex[$key] = $i
        if ($catExtIdHeader) {
            $uuid = $row.$catExtIdHeader
            if ($uuid) { $catUuidMap[$key] = [string]$uuid }
        }
    }
}

# We'll modify workbook only if not dry-run
$package = $null
$wsToUpdate = $null
if (-not $DryRun) {
    try {
        $package = Open-ExcelPackage -Path $Workbook -ErrorAction Stop
        $wsToUpdate = $package.Workbook.Worksheets['ToUpdate']
        if (-not $wsToUpdate) { throw "Worksheet 'ToUpdate' missing in live package (race?)" }
    } catch { Write-Error "Failed to open workbook for editing: $($_.Exception.Message)"; exit 2 }
}

# Helper: ensure header exists (returns column index)
function Ensure-Header {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][string]$HeaderText
    )
    $dim = $Worksheet.Dimension
    $maxCol = if ($dim) { $dim.End.Column } else { 0 }
    # Search existing headers row 1
    for ($c=1; $c -le $maxCol; $c++) {
        if (($Worksheet.Cells[1,$c].Text).Trim() -eq $HeaderText) { return $c }
    }
    $newCol = $maxCol + 1
    $Worksheet.Cells[1,$newCol].Value = $HeaderText
    return $newCol
}

$colCategoryUUIDs = $null
$colStatus        = $null
if ($wsToUpdate) {
    $colCategoryUUIDs = Ensure-Header -Worksheet $wsToUpdate -HeaderText 'Category UUID(s)'
    $colStatus        = Ensure-Header -Worksheet $wsToUpdate -HeaderText 'VM Name/extId & Category exId(s) Match'
}

# Styles
$greenColor = [System.Drawing.Color]::FromArgb(0x4C,0xAF,0x50)
$redColor   = [System.Drawing.Color]::FromArgb(0xE5,0x73,0x73)

function Apply-StatusStyle {
    param($Cell,[bool]$Success)
    if (-not $Cell) { return }
    $Cell.Style.Font.Bold = $true
    $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $Cell.Style.Fill.PatternType = 'Solid'
    if ($Success) { $Cell.Style.Fill.BackgroundColor.SetColor($greenColor) } else { $Cell.Style.Fill.BackgroundColor.SetColor($redColor) }
}

$anyMismatch = $false

for ($i=0; $i -lt $toUpdateData.Count; $i++) {
    $row = $toUpdateData[$i]
    $vmName  = $row.'VM Name'
    $vmExtId = $row.'VM extId'
    $spec    = $row.'UPDATE WITH CATEGORIES'
    if ([string]::IsNullOrWhiteSpace($vmName) -or [string]::IsNullOrWhiteSpace($vmExtId)) { Write-Verbose2 "Skipping blank VM row index $i" 2; continue }

    $vmKey = (Normalize $vmName)+'||'+(Normalize $vmExtId)
    $vmCount = $vmPairCounts[$vmKey]
    $vmMatchOk = ($vmCount -eq 1)

    $catPairs = @()
    $catValidOk = $true
    $spec = [string]$spec
    if (-not [string]::IsNullOrWhiteSpace($spec)) {
        $frags = $spec.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
        foreach ($frag in $frags) {
            if ($frag -notmatch '=') { Write-Host "Mismatch: VM '$vmName' extId '$vmExtId' category fragment missing '=': '$frag'"; $catValidOk = $false; continue }
            $parts = $frag.Split('=',2)
            $c = $parts[0].Trim(); $v = $parts[1].Trim()
            $key = (Normalize $c)+'||'+(Normalize $v)
            $catPairs += [PSCustomObject]@{ Cat=$c; Val=$v; Key=$key }
            if (-not $catRowIndex.ContainsKey($key)) { Write-Host "Mismatch: Category='$c' Value='$v' not found for VM '$vmName' extId '$vmExtId'"; $catValidOk = $false }
        }
    } else {
        Write-Host "Mismatch: No UPDATE WITH CATEGORIES specified for VM '$vmName' extId '$vmExtId'"; $catValidOk = $false
    }

    $ok = $vmMatchOk -and $catValidOk

    if ($ok) {
        $uuids = @()
        if ($catExtIdHeader) {
            foreach ($p in $catPairs) {
                if ($catUuidMap.ContainsKey($p.Key)) {
                    $uuid = $catUuidMap[$p.Key]
                    if ($uuid -and ($uuids -notcontains $uuid)) { $uuids += $uuid; Write-Verbose2 "Resolved $($p.Cat)=$($p.Val) -> $uuid" 2 }
                } else {
                    Write-Verbose2 "No UUID for $($p.Cat)=$($p.Val) (missing extID)" 2
                }
            }
        }
        $uuidList = ($uuids -join ',')
        if (-not $DryRun -and $wsToUpdate) {
            $excelRow = $i + 2  # data row offset (+ header)
            if ($colCategoryUUIDs) { $wsToUpdate.Cells[$excelRow,$colCategoryUUIDs].Value = $uuidList }
            if ($colStatus) { $cell = $wsToUpdate.Cells[$excelRow,$colStatus]; $cell.Value = 'OK'; Apply-StatusStyle -Cell $cell -Success $true }
        }
        Write-Host "OK, all matches for VM '$vmName' extId '$vmExtId'"
    }
    else {
        $anyMismatch = $true
        if (-not $vmMatchOk) {
            if (-not $vmCount) { Write-Host "Mismatch: VM Name/extId pair not found in VMCategories: '$vmName' / '$vmExtId'" }
            else { Write-Host "Mismatch: Duplicate VM Name/extId pair count=$vmCount in VMCategories: '$vmName' / '$vmExtId'" }
        }
        # category mismatches already printed
        if (-not $DryRun -and $wsToUpdate) {
            $excelRow = $i + 2
            if ($colStatus) { $cell = $wsToUpdate.Cells[$excelRow,$colStatus]; $cell.Value = 'Mismatch'; Apply-StatusStyle -Cell $cell -Success $false }
        }
    }
}

if (-not $DryRun -and $package) {
    try { Close-ExcelPackage $package } catch { Write-Error "Failed to save workbook: $($_.Exception.Message)"; exit 2 }
} elseif ($DryRun) {
    Write-Host "DryRun: no changes written." -ForegroundColor Yellow
}

if ($anyMismatch) { exit 1 } else { exit 0 }
