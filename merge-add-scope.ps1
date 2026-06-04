# merge-add-scope.ps1
# One-shot migration: reads sharepoint-source.xlsx (a copy of the live
# SharePoint workbook), inserts a `Scope` column right after `Future`, and
# writes GPS-Ecosystem.xlsx. Every other column is preserved verbatim from
# the source.
#
# After this runs once, the SharePoint workbook is authoritative going
# forward — do NOT run build-xlsx.ps1 against an edited workbook again,
# or editor changes will be stomped.
#
# Usage:
#   PowerShell -ExecutionPolicy Bypass -File .\merge-add-scope.ps1

$ErrorActionPreference = 'Stop'

$src  = Join-Path $PSScriptRoot 'sharepoint-source.xlsx'
$dst  = Join-Path $PSScriptRoot 'GPS-Ecosystem.xlsx'
if (-not (Test-Path $src)) { throw "Source not found: $src" }

# Items that should appear in the "Essential" toggle mode by default.
# Editors can change any cell in the Scope column afterward — this just
# seeds reasonable starting values. Names below must match the Name column
# in the SharePoint workbook exactly.
$essentialNames = @(
    'Educators & Teachers', 'School & District Leaders', 'SCDE Staff', 'Public',
    'Teacher Navigator', 'Admin Navigator',
    'School Report Cards',
    'Classlink',
    'PowerSchool SIS', 'Students Assessments',
    'Stadium Data Warehouse', 'Ed-Fi Integration Layer', 'Data Standards & APIs', 'Security & Governance',
    'GPS Leadership', 'X-Team', 'SCDE Offices'
)

$scopeOptions = 'Essential,Current,Future'

Write-Host "Reading $src ..."
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open a copy (don't modify the source in place; we'll Save As)
    $wb = $excel.Workbooks.Open($src)
    $ws = $wb.Worksheets.Item('Items')

    # Locate the Future column by header (defensive — column letter may shift
    # if editors added columns in SharePoint).
    $usedRange = $ws.UsedRange
    $colCount = $usedRange.Columns.Count
    $rowCount = $usedRange.Rows.Count

    $futureCol = $null
    $nameCol   = $null
    for ($c = 1; $c -le $colCount; $c++) {
        $h = $ws.Cells.Item(1, $c).Value2
        if ($h -eq 'Future') { $futureCol = $c }
        if ($h -eq 'Name')   { $nameCol = $c }
    }
    if (-not $futureCol) { throw "No 'Future' column found in source workbook." }
    if (-not $nameCol)   { throw "No 'Name' column found in source workbook." }
    Write-Host "Future column at index $futureCol; Name column at index $nameCol; total cols = $colCount; rows = $rowCount"

    # If a Scope column already exists, abort — caller can drop and retry.
    for ($c = 1; $c -le $colCount; $c++) {
        if ($ws.Cells.Item(1, $c).Value2 -eq 'Scope') {
            throw "A 'Scope' column already exists in the source. Remove it manually before re-running this script."
        }
    }

    # Insert a new column at $futureCol + 1
    $scopeCol = $futureCol + 1
    $insertRange = $ws.Columns.Item($scopeCol)
    [void]$insertRange.Insert(-4161, 0)  # xlShiftToRight = -4161, xlFormatFromLeftOrAbove = 0

    # Header
    $ws.Cells.Item(1, $scopeCol).Value2 = 'Scope'
    $headerCell = $ws.Cells.Item(1, $scopeCol)
    $headerCell.Font.Bold = $true
    $headerCell.Interior.Color = 3092271  # SCDE slate
    $headerCell.Font.Color = 16777215     # white
    $ws.Columns.Item($scopeCol).ColumnWidth = 11

    # Fill each data row's Scope based on the rules:
    #   - Name in $essentialNames    -> 'Essential'
    #   - Future cell == TRUE        -> 'Future'
    #   - otherwise                  -> 'Current'
    for ($r = 2; $r -le $rowCount; $r++) {
        $name      = $ws.Cells.Item($r, $nameCol).Value2
        $futureVal = $ws.Cells.Item($r, $futureCol).Value2
        $isFuture  = $false
        if ($futureVal -is [bool]) { $isFuture = $futureVal }
        elseif ($futureVal -is [string]) { $isFuture = ($futureVal.Trim().ToUpper() -eq 'TRUE') }

        $scope = 'Current'
        if ($essentialNames -contains $name) { $scope = 'Essential' }
        elseif ($isFuture) { $scope = 'Future' }

        $ws.Cells.Item($r, $scopeCol).Value2 = $scope
    }

    # Data validation dropdown for the new column
    $col_letter = if ($scopeCol -le 26) { [char]([byte][char]'A' + $scopeCol - 1) }
                  else { throw "Scope column index $scopeCol beyond column Z; extend this script if needed." }
    $scopeRange = $ws.Range("$($col_letter)2:$($col_letter)500")
    $scopeRange.Validation.Delete()
    [void]$scopeRange.Validation.Add(3, 1, 1, $scopeOptions)
    $scopeRange.Validation.IgnoreBlank = $true
    $scopeRange.Validation.InCellDropdown = $true
    $scopeRange.Validation.ErrorTitle = 'Invalid scope'
    $scopeRange.Validation.ErrorMessage = "Pick one of: $scopeOptions"

    # Save As GPS-Ecosystem.xlsx (preserves all other formatting, validation, etc.)
    if (Test-Path $dst) { Remove-Item $dst -Force }
    $wb.SaveAs($dst, 51)  # xlOpenXMLWorkbook = 51
    $wb.Close($false)
    Write-Host "Wrote $dst"
} finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect() | Out-Null
    [System.GC]::WaitForPendingFinalizers() | Out-Null
}
