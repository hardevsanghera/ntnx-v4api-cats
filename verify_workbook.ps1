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
  [string]$Path = "$PWD\scratch\cat_map.xlsx",
  [string]$WorksheetName = 'VMCategories',
  [int]$Head = 10,
  [switch]$Validate
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

if (-not (Test-Path $Path)) { Write-Err "Workbook not found: $Path"; exit 1 }

try {
  $pkg = Open-ExcelPackage -Path $Path -ErrorAction Stop
} catch {
  Write-Err "Failed to open workbook: $($_.Exception.Message)"; exit 1
}

$ws = $pkg.Workbook.Worksheets[$WorksheetName]
if (-not $ws) {
  $names = $pkg.Workbook.Worksheets | ForEach-Object { $_.Name }
  Write-Err "Worksheet '$WorksheetName' not found. Available: $(($names -join ', '))"; exit 1
}

if (-not $ws.Dimension) { Write-Warn "Worksheet has no data."; exit 0 }

$endRow = $ws.Dimension.End.Row
$endCol = $ws.Dimension.End.Column

# Read header row text preserving exact case
$headers = for ($c = 1; $c -le $endCol; $c++) { [string]$ws.Cells[1,$c].Text }

Write-Info "Headers ($($headers.Count))"
for ($i = 0; $i -lt $headers.Count; $i++) {
  Write-Host ("[{0,2}] {1}" -f ($i+1), $headers[$i])
}

# Print first N data rows as delimited text
$maxRow = [Math]::Min($endRow, $Head + 1)
if ($maxRow -le 1) { Write-Warn "No data rows present."; exit 0 }

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
  } else {
    Write-Warn "'VM Name' column not found."
  }
}

try { Close-ExcelPackage $pkg } catch { }

Write-Host "\nDone." -ForegroundColor Green
