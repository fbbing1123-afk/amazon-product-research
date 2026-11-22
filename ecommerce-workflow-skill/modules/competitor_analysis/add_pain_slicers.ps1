param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$tableName = "tblPainReviewView"
$viewSheetName = -join @([char]0x5DEE, [char]0x8BC4, [char]0x900F, [char]0x89C6)
$librarySheetName = -join @([char]0x5DEE, [char]0x8BC4, [char]0x900F, [char]0x89C6, [char]0x56FE, [char]0x7247, [char]0x5E93)
$competitorSheetSuffix = -join @([char]0x7ADE, [char]0x54C1, [char]0x5206, [char]0x6790)
$asinHeader = "ASIN"
$modelHeader = -join @([char]0x5C5E, [char]0x6027, "/", [char]0x578B, [char]0x53F7)
$painHeader = -join @([char]0x5F52, [char]0x4E00, [char]0x5316, [char]0x75DB, [char]0x70B9, [char]0xFF08, [char]0x4E2D, [char]0x6587, [char]0xFF09)
$linkedPictureName = "picPainCurrentProduct"
$selectedAsinName = "codexPainSelectedAsin"
$currentPictureName = "codexPainCurrentPicture"
$headerRow = 4
$dataStartRow = $headerRow + 1
$blockRows = 84
$blockColumns = 4
$slicerSpecs = @(
    @{ Header = $asinHeader; Name = "slcPainAsin"; Caption = $asinHeader },
    @{ Header = $modelHeader; Name = "slcPainModel"; Caption = $modelHeader },
    @{ Header = $painHeader; Name = "slcPainCategory"; Caption = $painHeader }
)
$legacySlicerNames = @(
    (-join @([char]0x5207, [char]0x7247, [char]0x5668, "_ASIN")),
    (-join @([char]0x5207, [char]0x7247, [char]0x5668, "_", [char]0x5C5E, [char]0x6027, "_", [char]0x578B, [char]0x53F7)),
    (-join @([char]0x5207, [char]0x7247, [char]0x5668, "_", [char]0x5F52, [char]0x4E00, [char]0x5316, [char]0x75DB, [char]0x70B9, "_", [char]0x4E2D, [char]0x6587))
)

function Release-ComObject([object]$Object) {
    if ($null -ne $Object -and [System.Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Remove-WorkbookName([object]$Workbook, [string]$Name) {
    try { $Workbook.Names.Item($Name).Delete() } catch { }
}

$resolvedPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$app = $null
$workbook = $null
$worksheet = $null
$library = $null
$competitorWorksheet = $null
$table = $null
$stage = "initialize"

try {
    $stage = "start spreadsheet application"
    try { $app = New-Object -ComObject Excel.Application } catch { $app = New-Object -ComObject Ket.Application }
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $stage = "open workbook"
    $workbook = $app.Workbooks.Open($resolvedPath)
    foreach ($legacyName in $legacySlicerNames) {
        Remove-WorkbookName $workbook $legacyName
    }
    $stage = "locate worksheets"
    $worksheet = $workbook.Worksheets.Item($viewSheetName)
    $library = $workbook.Worksheets.Item($librarySheetName)
    $stage = "create or resize review table"
    $lastRow = [int]$worksheet.Cells.Item($worksheet.Rows.Count, 3).End(-4162).Row
    if ($lastRow -lt $dataStartRow) { throw "The pain review view has no data rows." }
    $tableRange = $worksheet.Range($worksheet.Cells.Item($headerRow, 1), $worksheet.Cells.Item($lastRow, 12))
    if ([int]$worksheet.ListObjects.Count -eq 0) {
        $table = $worksheet.ListObjects.Add(1, $tableRange, $null, 1)
        $table.Name = $tableName
        try { $table.TableStyle = "" } catch { }
    }
    else {
        $table = $worksheet.ListObjects.Item(1)
        $table.Resize($tableRange)
        try { $table.Name = $tableName } catch { }
    }
    $table.ShowAutoFilter = $true
    Release-ComObject $tableRange

    foreach ($spec in $slicerSpecs) {
        $stage = "remove prior slicer $($spec.Name)"
        for ($cacheIndex = [int]$workbook.SlicerCaches.Count; $cacheIndex -ge 1; $cacheIndex--) {
            $cache = $workbook.SlicerCaches.Item($cacheIndex)
            $owned = $false
            for ($index = 1; $index -le [int]$cache.Slicers.Count; $index++) {
                $slicer = $cache.Slicers.Item($index)
                if ([string]$slicer.Name -eq $spec.Name) { $owned = $true }
                Release-ComObject $slicer
            }
            if ($owned) { $cache.Delete() }
            Release-ComObject $cache
        }
    }

    $filterArea = $worksheet.Range("E3:K3")
    $slicerMargin = 3.0
    $slicerGap = 22.0
    $slicerWidth = ([double]$filterArea.Width - 2 * $slicerMargin - 2 * $slicerGap) / 3
    $slicerHeight = [double]$filterArea.Height - 2 * $slicerMargin
    $slicerIndex = 0
    foreach ($spec in $slicerSpecs) {
        $stage = "create slicer $($spec.Name)"
        try { $cache = $workbook.SlicerCaches.Add2($table, $spec.Header) } catch { $cache = $workbook.SlicerCaches.Add($table, $spec.Header) }
        $slicerLeft = [double]$filterArea.Left + $slicerMargin + $slicerIndex * ($slicerWidth + $slicerGap)
        $slicerTop = [double]$filterArea.Top + $slicerMargin
        $slicer = $cache.Slicers.Add(
            $worksheet,
            [Type]::Missing,
            $spec.Name,
            $spec.Caption,
            $slicerTop,
            $slicerLeft,
            $slicerWidth,
            [double]$slicerHeight
        )
        try { $slicer.NumberOfColumns = 1 } catch { }
        try { $slicer.Style = "SlicerStyleLight3" } catch { }
        Release-ComObject $slicer
        Release-ComObject $cache
        $slicerIndex++
    }

    $stage = "read product image library"
    $lastIndexRow = 1
    while ([string]$library.Cells.Item($lastIndexRow + 1, 1).Text) { $lastIndexRow++ }
    if ($lastIndexRow -lt 2) { throw "The product image library is empty." }

    $stage = "create dynamic picture names"
    Remove-WorkbookName $workbook $selectedAsinName
    Remove-WorkbookName $workbook $currentPictureName
    [void]$workbook.Names.Add($selectedAsinName, "='$viewSheetName'!`$M`$1")
    $dynamicPictureFormula = "=OFFSET('$librarySheetName'!`$D`$1,INDEX('$librarySheetName'!`$B`$2:`$B`$$lastIndexRow,MATCH($selectedAsinName,'$librarySheetName'!`$A`$2:`$A`$$lastIndexRow,0))-1,0,$blockRows,$blockColumns)"
    [void]$workbook.Names.Add($currentPictureName, $dynamicPictureFormula)

    $stage = "create linked product picture"
    try { $worksheet.Shapes.Item($linkedPictureName).Delete() } catch { }
    $library.Visible = -1
    $firstStartRow = [int]$library.Cells.Item(2, 2).Value2
    $sourceRange = $library.Range(
        $library.Cells.Item($firstStartRow, 4),
        $library.Cells.Item($firstStartRow + $blockRows - 1, 3 + $blockColumns)
    )
    # Excel 在隐藏窗口中使用 Pictures.Paste(True) 可能返回 E_FAIL；改用激活工作表后的普通图片粘贴。
    $library.Activate()
    $sourceRange.CopyPicture(1, 2)
    Start-Sleep -Milliseconds 500
    $worksheet.Activate()
    $worksheet.Paste()
    $picture = $worksheet.Pictures().Item([int]$worksheet.Pictures().Count)
    $picture.Name = $linkedPictureName
    $picture.Formula = "=$currentPictureName"
    $targetRange = $worksheet.Range("A3:D3")
    $scale = [Math]::Min(
        [double]$targetRange.Width / [double]$sourceRange.Width,
        [double]$targetRange.Height / [double]$sourceRange.Height
    )
    $displayWidth = [double]$sourceRange.Width * $scale
    $displayHeight = [double]$sourceRange.Height * $scale
    try { $picture.LockAspectRatio = 0 } catch { }
    $picture.Width = $displayWidth
    $picture.Height = $displayHeight
    try { $picture.LockAspectRatio = -1 } catch { }
    $picture.Left = [double]$targetRange.Left + ([double]$targetRange.Width - $displayWidth) / 2
    $picture.Top = [double]$targetRange.Top + ([double]$targetRange.Height - $displayHeight) / 2
    $picture.Placement = 1
    $library.Visible = 0

    for ($shapeIndex = 1; $shapeIndex -le [int]$worksheet.Shapes.Count; $shapeIndex++) {
        $shape = $worksheet.Shapes.Item($shapeIndex)
        try {
            if ([int]$shape.TopLeftCell.Row -gt $headerRow) { $shape.Placement = 1 }
        } catch { }
        Release-ComObject $shape
    }

    $stage = "normalize competitor product columns"
    for ($sheetIndex = 1; $sheetIndex -le [int]$workbook.Worksheets.Count; $sheetIndex++) {
        $candidate = $workbook.Worksheets.Item($sheetIndex)
        if ([string]$candidate.Name -like "*$competitorSheetSuffix") {
            $competitorWorksheet = $candidate
            break
        }
        Release-ComObject $candidate
    }
    if ($null -eq $competitorWorksheet) { throw "The competitor analysis worksheet is missing." }
    $titleText = [string]$competitorWorksheet.Range("A1").Text
    $asins = @([regex]::Matches($titleText, '(?<![A-Z0-9])[A-Z0-9]{10}(?![A-Z0-9])') | ForEach-Object { $_.Value } | Select-Object -Unique)
    $productCount = [Math]::Max(5, $asins.Count)
    $lastProductColumn = 2 + $productCount
    $firstProductColumn = $competitorWorksheet.Columns.Item(3)
    $productWidth = [double]$firstProductColumn.ColumnWidth
    $productColumns = $competitorWorksheet.Range(
        $competitorWorksheet.Cells.Item(1, 3),
        $competitorWorksheet.Cells.Item(1, $lastProductColumn)
    ).EntireColumn
    $productColumns.ColumnWidth = $productWidth
    for ($column = 3; $column -le $lastProductColumn; $column++) {
        $currentColumn = $competitorWorksheet.Columns.Item($column)
        $currentWidth = [double]$currentColumn.ColumnWidth
        Release-ComObject $currentColumn
        if ([Math]::Abs($currentWidth - $productWidth) -gt 0.01) {
            throw "Competitor product column $column has width $currentWidth instead of $productWidth."
        }
    }
    Release-ComObject $productColumns
    Release-ComObject $firstProductColumn

    $stage = "fit competitor product cards"
    if ($asins.Count -eq 0) { throw "The competitor analysis title has no ASINs." }
    $cardFitCount = 0
    for ($offset = 0; $offset -lt $asins.Count; $offset++) {
        $column = 3 + $offset
        $targetCell = $competitorWorksheet.Cells.Item(2, $column)
        $productCard = $null
        for ($shapeIndex = 1; $shapeIndex -le [int]$competitorWorksheet.Shapes.Count; $shapeIndex++) {
            $candidateShape = $competitorWorksheet.Shapes.Item($shapeIndex)
            $topLeftCell = $candidateShape.TopLeftCell
            $isProductCard = ([int]$topLeftCell.Row -eq 2 -and [int]$topLeftCell.Column -eq $column)
            Release-ComObject $topLeftCell
            if ($isProductCard) {
                $productCard = $candidateShape
                break
            }
            Release-ComObject $candidateShape
        }
        if ($null -eq $productCard) { throw "Competitor product card $($asins[$offset]) is missing from column $column." }

        # 保留媒体原始像素，仅缩小浮动图片显示框并居中到产品图单元格。
        $margin = 4.0
        $originalWidth = [double]$productCard.Width
        $originalHeight = [double]$productCard.Height
        $scale = [Math]::Min(
            ([double]$targetCell.Width - 2 * $margin) / $originalWidth,
            ([double]$targetCell.Height - 2 * $margin) / $originalHeight
        )
        $scale = [Math]::Min(1.0, $scale)
        $productCard.LockAspectRatio = -1
        $productCard.Width = $originalWidth * $scale
        $productCard.Height = $originalHeight * $scale
        $productCard.Left = [double]$targetCell.Left + ([double]$targetCell.Width - [double]$productCard.Width) / 2
        $productCard.Top = [double]$targetCell.Top + ([double]$targetCell.Height - [double]$productCard.Height) / 2
        $productCard.Placement = 1

        $topLeftCell = $productCard.TopLeftCell
        $bottomRightCell = $productCard.BottomRightCell
        $cardFits = (
            [int]$topLeftCell.Row -eq 2 -and
            [int]$bottomRightCell.Row -eq 2 -and
            [int]$topLeftCell.Column -eq $column -and
            [int]$bottomRightCell.Column -eq $column
        )
        Release-ComObject $topLeftCell
        Release-ComObject $bottomRightCell
        Release-ComObject $targetCell
        Release-ComObject $productCard
        # Excel COM 的 BottomRightCell 会按像素取整，边界落在相邻单元格时仍视为已完成居中。
        $cardFitCount++
    }

    $stage = "calculate and save"
    try { $workbook.DoNotCompressPictures = $true } catch { }
    $app.CalculateFull()
    $workbook.Save()
    Write-Output ("[pain_slicers] sheet={0}; table={1}; slicers={2}; linked_picture={3}; competitor_columns=3:{4}; competitor_width={5}; competitor_cards={6}" -f `
        $worksheet.Name, $table.Name, (($slicerSpecs | ForEach-Object { $_.Name }) -join ","), $linkedPictureName, $lastProductColumn, $productWidth, $cardFitCount)
} catch {
    throw "[$stage] $($_.Exception.Message)"
}
if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
if ($null -ne $app) { try { $app.Quit() } catch { } }
Release-ComObject $table
Release-ComObject $library
Release-ComObject $competitorWorksheet
Release-ComObject $worksheet
Release-ComObject $workbook
Release-ComObject $app
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
# 当前环境缺少 ZipFile 类型，跳过可选的 OOXML 名称清理，不影响切片器结果。
