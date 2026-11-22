param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$tableName = "tblMarketSlicerView"
$viewSheetName = -join @([char]0x5207, [char]0x7247, [char]0x8054, [char]0x52A8)
$headerRow = 12
$salesLayerHeader = -join @([char]0x9500, [char]0x91CF, [char]0x5206, [char]0x5C42)
$reviewBucketHeader = -join @([char]0x8BC4, [char]0x5206, [char]0x6570, [char]0x533A, [char]0x95F4)
$priceBucketHeader = -join @([char]0x4EF7, [char]0x683C, [char]0x6BB5)
$slicerSpecs = @(
    @{ Header = $salesLayerHeader; Name = "slcMarketSalesLayer"; Caption = $salesLayerHeader; Column = 2; Width = 190; Height = 175 },
    @{ Header = $reviewBucketHeader; Name = "slcMarketReviewBucket"; Caption = $reviewBucketHeader; Column = 6; Width = 190; Height = 195 },
    @{ Header = $priceBucketHeader; Name = "slcMarketPriceBucket"; Caption = $priceBucketHeader; Column = 10; Width = 190; Height = 185 }
)

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

$resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$app = $null
$workbook = $null
$worksheet = $null

try {
    try {
        $app = New-Object -ComObject Excel.Application
    }
    catch {
        $app = New-Object -ComObject Ket.Application
    }
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $workbook = $app.Workbooks.Open($resolvedPath)

    foreach ($candidate in @($workbook.Worksheets)) {
        if ([string]$candidate.Name -eq $viewSheetName) {
            $worksheet = $candidate
            break
        }
    }
    if ($null -eq $worksheet) {
        throw "Slicer view worksheet was not found."
    }

    $usedRange = $worksheet.UsedRange
    $firstRow = [int]$usedRange.Row
    $firstColumn = [int]$usedRange.Column
    $lastRow = $firstRow + [int]$usedRange.Rows.Count - 1
    $lastColumn = $firstColumn + [int]$usedRange.Columns.Count - 1
    if ($firstRow -ne 1 -or $lastRow -le $headerRow) {
        throw "The slicer view worksheet has no data rows."
    }

    $headerColumns = @{}
    for ($column = 1; $column -le $lastColumn; $column++) {
        $header = [string]$worksheet.Cells.Item($headerRow, $column).Text
        if ($header) { $headerColumns[$header.Trim()] = $column }
    }
    foreach ($spec in $slicerSpecs) {
        if (-not $headerColumns.ContainsKey($spec.Header)) {
            throw "The source worksheet is missing a slicer helper column: $($spec.Header)"
        }
    }

    $firstHelperColumn = ($slicerSpecs | ForEach-Object { [int]$headerColumns[$_.Header] } | Measure-Object -Minimum).Minimum
    $sourceValueRange = $worksheet.Range(
        $worksheet.Cells.Item($headerRow + 1, 1),
        $worksheet.Cells.Item($lastRow, $firstHelperColumn - 1)
    )
    $sourceValues = $sourceValueRange.Value2
    $sourceColumnCount = $firstHelperColumn - 1
    $sourceRowCount = $lastRow - $headerRow
    $lastDataRow = $headerRow
    for ($rowOffset = $sourceRowCount; $rowOffset -ge 1 -and $lastDataRow -eq $headerRow; $rowOffset--) {
        for ($columnOffset = 1; $columnOffset -le $sourceColumnCount; $columnOffset++) {
            $value = $sourceValues[$rowOffset, $columnOffset]
            if ($null -ne $value -and [string]$value -ne "") {
                $lastDataRow = $rowOffset + $headerRow
                break
            }
        }
    }
    Release-ComObject $sourceValueRange
    if ($lastDataRow -le $headerRow) {
        throw "The source worksheet has no non-empty data rows."
    }
    $lastRow = $lastDataRow

    # Remove only slicers created by this module. Deleting their caches prevents duplicate controls on rerun.
    for ($cacheIndex = [int]$workbook.SlicerCaches.Count; $cacheIndex -ge 1; $cacheIndex--) {
        $cache = $workbook.SlicerCaches.Item($cacheIndex)
        $owned = $false
        for ($slicerIndex = 1; $slicerIndex -le [int]$cache.Slicers.Count; $slicerIndex++) {
            $slicer = $cache.Slicers.Item($slicerIndex)
            if ($slicerSpecs.Name -contains [string]$slicer.Name) { $owned = $true }
            Release-ComObject $slicer
        }
        if ($owned) { $cache.Delete() }
        Release-ComObject $cache
    }

    $table = $null
    $createdTable = $false
    for ($index = 1; $index -le [int]$worksheet.ListObjects.Count; $index++) {
        $candidateTable = $worksheet.ListObjects.Item($index)
        if ([string]$candidateTable.Name -eq $tableName) {
            $table = $candidateTable
            break
        }
        if ([int]$candidateTable.Range.Row -eq $headerRow -and [int]$candidateTable.Range.Column -eq 1) {
            $table = $candidateTable
            $tableName = [string]$candidateTable.Name
            break
        }
        Release-ComObject $candidateTable
    }
    $tableRange = $worksheet.Range($worksheet.Cells.Item($headerRow, 1), $worksheet.Cells.Item($lastRow, $lastColumn))
    if ($null -eq $table) {
        $table = $worksheet.ListObjects.Add(1, $tableRange, $null, 1)
        $table.Name = $tableName
        $createdTable = $true
    }
    else {
        $table.Resize($tableRange)
    }
    $table.ShowAutoFilter = $true
    if ($createdTable) { try { $table.TableStyle = "" } catch { } }

    # Product pictures are floating shapes in the workflow. Move-and-size placement makes row filters hide them.
    for ($shapeIndex = 1; $shapeIndex -le [int]$worksheet.Shapes.Count; $shapeIndex++) {
        $shape = $worksheet.Shapes.Item($shapeIndex)
        try {
            $shapeRow = [int]$shape.TopLeftCell.Row
            if ($shapeRow -gt $headerRow -and $shapeRow -le $lastRow) { $shape.Placement = 1 }
        }
        catch { }
        Release-ComObject $shape
    }

    $created = @()
    foreach ($spec in $slicerSpecs) {
        try {
            $cache = $workbook.SlicerCaches.Add2($table, $spec.Header)
        }
        catch {
            $cache = $workbook.SlicerCaches.Add($table, $spec.Header)
        }
        $slicer = $cache.Slicers.Add(
            $worksheet,
            [Type]::Missing,
            $spec.Name,
            $spec.Caption,
            [double]$worksheet.Cells.Item(2, $spec.Column).Top,
            [double]$worksheet.Cells.Item(2, $spec.Column).Left,
            [double]$spec.Width,
            [double]$spec.Height
        )
        $created += $spec.Name
        Release-ComObject $slicer
        Release-ComObject $cache
    }

    $workbook.Save()
    Write-Output ("[source_slicers] sheet={0}; rows={1}; columns={2}; table={3}; slicers={4}" -f `
        $worksheet.Name, ($lastRow - $headerRow), $lastColumn, $tableName, ($created -join ","))
}
finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
    if ($null -ne $app) { try { $app.Quit() } catch { } }
    Release-ComObject $worksheet
    Release-ComObject $workbook
    Release-ComObject $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
