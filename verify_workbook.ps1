<#
.SYNOPSIS
  Quick verifier for the generated workbook (scratch\cat_map.xlsx).

.DESCRIPTION
  - Shows header names with column indexes (preserving case)
  - Prints the first few data rows as raw cell text, avoiding property name collisions
  - Optionally validates presence of case-variant headers (Environment vs environment)
    and that VM rows like hard-vm-1.. are present

.PARAMETER Path
  Workbook path. Defaults to $PWD\scratch\cat_map.xlsx

.PARAMETER WorksheetName
  Worksheet name to inspect. Defaults to 'VMCategories'

.PARAMETER Head
  Number of data rows (after header) to show. Defaults to 10

.PARAMETER Validate
  When set, performs simple checks (case-variant headers and hard-vm-* presence)
#>

[CmdletBinding()]
param(
  # Inspect mode (existing behavior)
  [string]$Path = "$PWD\scratch\cat_map.xlsx",
  [string]$WorksheetName = 'VMCategories',
  [int]$Head = 10,
  [switch]$Validate,

  # Update mode (new): verify against available categories and update the target workbook
  [switch]$Update,
  [string]$Workbook = "$PWD\scratch\VMsToUpdate-PROD.xlsx",
  [switch]$IgnoreCase,
  [switch]$WriteUUIDs
)

function Write-Info($msg) { Write-Host $msg -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host $msg -ForegroundColor Yellow }
function Write-Err($msg)  { Write-Host $msg -ForegroundColor Red }

try {
  if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Err "ImportExcel module not installed. Install with: Install-Module ImportExcel -Scope CurrentUser"
    exit 2
  }
  Import-Module ImportExcel -ErrorAction Stop
} catch {
  Write-Err "Failed to import ImportExcel: $($_.Exception.Message)"; exit 2
}

# --- Helper: Normalize for case-sensitive/insensitive compare ---
function Normalize {
  param([object]$Value)
  if ($null -eq $Value) { return '' }
  $s = [string]$Value
  if ($IgnoreCase) { return $s.ToLowerInvariant() } else { return $s }
}

# --- Helper: read a sheet as headers + row arrays (preserve case) ---
function Read-WorksheetRows {
  param([string]$Path,[string]$Worksheet)
  try { $pkgLocal = Open-ExcelPackage -Path $Path -ErrorAction Stop } catch { throw }
  try {
    $wsLocal = $pkgLocal.Workbook.Worksheets[$Worksheet]
    if (-not $wsLocal) { throw "Worksheet not found: $Worksheet" }
    if (-not $wsLocal.Dimension) { return @{ Headers=@(); Rows=@() } }
    $endRowL = $wsLocal.Dimension.End.Row
    $endColL = $wsLocal.Dimension.End.Column
    $headersL = for ($c=1;$c -le $endColL;$c++){ [string]$wsLocal.Cells[1,$c].Text }
    $rowsL = @()
    for ($r=2;$r -le $endRowL;$r++) {
      $valsL = New-Object string[] $endColL
      for ($c=1;$c -le $endColL;$c++){ $valsL[$c-1] = [string]$wsLocal.Cells[$r,$c].Text }
      $rowsL += ,$valsL
    }
    return @{ Headers=$headersL; Rows=$rowsL }
  } finally { try { if ($pkgLocal) { Close-ExcelPackage $pkgLocal } } catch { } }
}

# --- Inspect mode (existing behavior) ---
if (-not $Update) {
  if (-not (Test-Path $Path)) { Write-Err "Workbook not found: $Path"; exit 1 }
  try { $pkg = Open-ExcelPackage -Path $Path -ErrorAction Stop }
  catch { Write-Err "Failed to open workbook: $($_.Exception.Message)"; exit 1 }

  $ws = $pkg.Workbook.Worksheets[$WorksheetName]
  if (-not $ws) {
    $names = $pkg.Workbook.Worksheets | ForEach-Object { $_.Name }
    Write-Err "Worksheet '$WorksheetName' not found. Available: $(($names -join ', '))"; exit 1
  }
  if (-not $ws.Dimension) { Write-Warn "Worksheet has no data."; try { Close-ExcelPackage $pkg } catch { }; exit 0 }

  $endRow = $ws.Dimension.End.Row
  $endCol = $ws.Dimension.End.Column
  $headers = for ($c = 1; $c -le $endCol; $c++) { [string]$ws.Cells[1,$c].Text }

  Write-Info "Headers ($($headers.Count))"
  for ($i = 0; $i -lt $headers.Count; $i++) { Write-Host ("[{0,2}] {1}" -f ($i+1), $headers[$i]) }

  $maxRow = [Math]::Min($endRow, $Head + 1)
  if ($maxRow -le 1) { Write-Warn "No data rows present."; try { Close-ExcelPackage $pkg } catch { }; exit 0 }

  Write-Info ("\nFirst {0} data row(s) (of {1})" -f ($maxRow-1), ($endRow-1))
  for ($r = 2; $r -le $maxRow; $r++) {
    $vals = for ($c = 1; $c -le $endCol; $c++) { [string]$ws.Cells[$r,$c].Text }
    Write-Host ("Row {0,4}: {1}" -f $r, ($vals -join ' | '))
  }

  if ($Validate) {
    $hasEnvCap   = $headers -contains 'Environment'
    $hasEnvLower = $headers -contains 'environment'
    if ($hasEnvCap -and $hasEnvLower) { Write-Info "Both 'Environment' and 'environment' headers present (case-distinct)." } else { Write-Warn "Case-variant headers not both present." }

    $idxVmName = [Array]::IndexOf($headers, 'VM Name')
    if ($idxVmName -ge 0) {
      $found = $false
      for ($r = 2; $r -le [Math]::Min($endRow, 50); $r++) {
        $vm = [string]$ws.Cells[$r, ($idxVmName+1)].Text
        if ($vm -match '^hard-vm-\d+$') { $found = $true; break }
      }
      if ($found) { Write-Info "Found VM rows like 'hard-vm-n'." } else { Write-Warn "Did not find 'hard-vm-n' rows in first 50 entries." }
    } else { Write-Warn "'VM Name' column not found." }
  }

  try { Close-ExcelPackage $pkg } catch { }
  Write-Host "\nDone." -ForegroundColor Green
  return
}

# --- Update mode: verify and update the target workbook ---
Write-Info "Update mode: verifying categories against available list and updating worksheet statuses..."
if (-not (Test-Path -LiteralPath $Workbook)) { Write-Err "Workbook not found: $Workbook"; exit 2 }

# Load ToUpdate via Import-Excel
try { $toUpdateData = Import-Excel -Path $Workbook -Worksheet 'ToUpdate' }
catch { Write-Err "Failed to read 'ToUpdate' sheet from ${Workbook}: $($_.Exception.Message)"; exit 2 }

# Load VMCategories + AllCategories via index-based reader (preserve header case)
try {
  $vmSheet  = Read-WorksheetRows -Path $Workbook -Worksheet 'VMCategories'
  $allSheet = Read-WorksheetRows -Path $Workbook -Worksheet 'AllCategories'
} catch { Write-Err "Failed to read required sheets from ${Workbook}: $($_.Exception.Message)"; exit 2 }

$idxVmName = [Array]::IndexOf($vmSheet.Headers,'VM Name')
$idxVmExt  = [Array]::IndexOf($vmSheet.Headers,'VM extId')
if ($idxVmName -lt 0 -or $idxVmExt -lt 0) { Write-Err "Missing 'VM Name' or 'VM extId' in VMCategories"; exit 2 }
$vmCatData = foreach ($row in $vmSheet.Rows) { [PSCustomObject]@{ 'VM Name' = $row[$idxVmName]; 'VM extId' = $row[$idxVmExt] } }

$idxCat    = [Array]::IndexOf($allSheet.Headers,'Category')
$idxVal    = [Array]::IndexOf($allSheet.Headers,'Value')
$idxExtID  = [Array]::IndexOf($allSheet.Headers,'extID'); if ($idxExtID -lt 0) { $idxExtID = [Array]::IndexOf($allSheet.Headers,'extId') }
if ($idxCat -lt 0 -or $idxVal -lt 0) { Write-Err "Missing 'Category' or 'Value' in AllCategories"; exit 2 }
$allCatData = foreach ($row in $allSheet.Rows) {
  $o = [ordered]@{ Category = $row[$idxCat]; Value = $row[$idxVal] }
  if ($idxExtID -ge 0) { $o['extID'] = $row[$idxExtID] }
  [PSCustomObject]$o
}

foreach ($c in 'VM Name','VM extId','UPDATE WITH CATEGORIES') { if (-not ($toUpdateData | Get-Member -Name $c -MemberType NoteProperty)) { Write-Err "Missing column '$c' in ToUpdate"; exit 2 } }

# Map VM pair counts and first row
$vmPairCounts = @{}
for ($i=0; $i -lt $vmCatData.Count; $i++) {
  $row = $vmCatData[$i]; $vmName=$row.'VM Name'; $vmExtId=$row.'VM extId'
  if ([string]::IsNullOrWhiteSpace($vmName) -and [string]::IsNullOrWhiteSpace($vmExtId)) { continue }
  $key = (Normalize $vmName)+'||'+(Normalize $vmExtId)
  $vmPairCounts[$key] = 1 + ($vmPairCounts[$key] | ForEach-Object { $_ })
}

# Map Category/Value to extID (if present)
$catRowIndex = @{}
$catUuidMap  = @{}
for ($i=0; $i -lt $allCatData.Count; $i++) {
  $row = $allCatData[$i]
  if ($null -eq $row) { continue }
  $cat = $row.Category; $val=$row.Value
  if ([string]::IsNullOrWhiteSpace($cat) -and [string]::IsNullOrWhiteSpace($val)) { continue }
  $key = (Normalize $cat)+'||'+(Normalize $val)
  if (-not $catRowIndex.ContainsKey($key)) {
    $catRowIndex[$key] = $i
    if ($row.PSObject.Properties.Match('extID')) { $uuid=$row.extID; if ($uuid) { $catUuidMap[$key]=[string]$uuid } }
  }
}

# Open workbook for editing 'ToUpdate'
$package = $null; $wsToUpdate = $null
try {
  $package = Open-ExcelPackage -Path $Workbook -ErrorAction Stop
  $wsToUpdate = $package.Workbook.Worksheets['ToUpdate']
  if (-not $wsToUpdate) { throw "Worksheet 'ToUpdate' missing in live package" }
} catch { Write-Err "Failed to open workbook for editing: $($_.Exception.Message)"; exit 2 }

function Ensure-Header {
  param($Worksheet,[string]$HeaderText)
  $dim = $Worksheet.Dimension; $maxCol = if ($dim) { $dim.End.Column } else { 0 }
  for ($c=1; $c -le $maxCol; $c++) { if (($Worksheet.Cells[1,$c].Text).Trim() -eq $HeaderText) { return $c } }
  $newCol = $maxCol + 1; $Worksheet.Cells[1,$newCol].Value = $HeaderText; return $newCol
}

$colStatus = Ensure-Header -Worksheet $wsToUpdate -HeaderText 'VM Name/extId & Category exId(s) Match'
$colUUIDs  = $null
if ($WriteUUIDs) { $colUUIDs = Ensure-Header -Worksheet $wsToUpdate -HeaderText 'Category UUID(s)' }

$greenColor = [System.Drawing.Color]::FromArgb(0x4C,0xAF,0x50)
$redColor   = [System.Drawing.Color]::FromArgb(0xE5,0x73,0x73)
function Apply-StatusStyle { param($Cell,[bool]$Success) if (-not $Cell) { return }; $Cell.Style.Font.Bold=$true; $Cell.Style.Font.Color.SetColor([System.Drawing.Color]::White); $Cell.Style.Fill.PatternType='Solid'; if ($Success) { $Cell.Style.Fill.BackgroundColor.SetColor($greenColor) } else { $Cell.Style.Fill.BackgroundColor.SetColor($redColor) } }

$anyMismatch = $false
for ($i=0; $i -lt $toUpdateData.Count; $i++) {
  $row = $toUpdateData[$i]
  $vmName=[string]$row.'VM Name'; $vmExtId=[string]$row.'VM extId'; $spec=[string]$row.'UPDATE WITH CATEGORIES'
  if ([string]::IsNullOrWhiteSpace($vmName) -or [string]::IsNullOrWhiteSpace($vmExtId)) { continue }
  $vmKey=(Normalize $vmName)+'||'+(Normalize $vmExtId); $vmCount=$vmPairCounts[$vmKey]
  $vmMatchOk = ($vmCount -eq 1)
  $catValidOk = $true; $catPairs=@()
  if (-not [string]::IsNullOrWhiteSpace($spec)) {
    $frags = $spec.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    foreach ($frag in $frags) {
      if ($frag -notmatch '=') { $catValidOk = $false; continue }
      $parts=$frag.Split('=',2); $c=$parts[0].Trim(); $v=$parts[1].Trim(); $key=(Normalize $c)+'||'+(Normalize $v); $catPairs += [PSCustomObject]@{ Key=$key; C=$c; V=$v }
      if (-not $catRowIndex.ContainsKey($key)) { $catValidOk = $false }
    }
  } else { $catValidOk = $false }

  $ok = $vmMatchOk -and $catValidOk
  $excelRow = $i + 2
  if ($ok) {
    if ($colStatus) { $cell=$wsToUpdate.Cells[$excelRow,$colStatus]; $cell.Value='OK'; Apply-StatusStyle -Cell $cell -Success $true }
    if ($WriteUUIDs -and $colUUIDs) {
      $uuids=@(); foreach ($p in $catPairs) { if ($catUuidMap.ContainsKey($p.Key)) { $u=$catUuidMap[$p.Key]; if ($u -and ($uuids -notcontains $u)) { $uuids += $u } } }
      $wsToUpdate.Cells[$excelRow,$colUUIDs].Value = ($uuids -join ',')
    }
  } else {
    $anyMismatch = $true
    if ($colStatus) { $cell=$wsToUpdate.Cells[$excelRow,$colStatus]; $cell.Value='Mismatch'; Apply-StatusStyle -Cell $cell -Success $false }
    if ($WriteUUIDs -and $colUUIDs) { $wsToUpdate.Cells[$excelRow,$colUUIDs].Value = '' }
  }
}

try { Close-ExcelPackage $package } catch { Write-Err "Failed to save workbook: $($_.Exception.Message)"; exit 2 }

if ($anyMismatch) { Write-Warn "Completed with mismatches. See 'VM Name/extId & Category exId(s) Match' column."; exit 1 }
Write-Host "Update complete. All rows OK." -ForegroundColor Green
exit 0
