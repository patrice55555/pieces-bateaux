param(
    [string]$InputCsv = ".\princess_parts_products.csv",
    [string]$OutputCsv = ".\princess_parts_products.csv",
    [string]$OutputXlsx = ".\princess_parts_products.xlsx"
)

$columnOrder = @(
    'category',
    'name',
    'description',
    'specifications',
    'model_number',
    'price_gbp_numeric',
    'stock',
    'delivery_delay',
    'url'
)

function Normalize-CellText([string]$Value) {
    if ($null -eq $Value) {
        return ''
    }

    return (($Value -replace "`r?`n", ' ') -replace '\s+', ' ').Trim()
}

function Set-ColumnWidthCm($worksheet, $excel, [int]$columnIndex, [double]$centimeters) {
    $targetWidth = $excel.CentimetersToPoints($centimeters)
    $column = $worksheet.Columns.Item($columnIndex)
    $low = 0.0
    $high = 100.0

    for ($i = 0; $i -lt 20; $i++) {
        $mid = ($low + $high) / 2.0
        $column.ColumnWidth = $mid

        if ($column.Width -lt $targetWidth) {
            $low = $mid
        }
        else {
            $high = $mid
        }
    }

    $column.ColumnWidth = [Math]::Round($high, 2)
}

$rows = Import-Csv $InputCsv

$cleanRows = foreach ($row in $rows) {
    [pscustomobject]@{
        category = Normalize-CellText $row.category
        name = Normalize-CellText $row.name
        description = Normalize-CellText $row.description
        specifications = Normalize-CellText $row.specifications
        model_number = Normalize-CellText $row.model_number
        price_gbp_numeric = Normalize-CellText $row.price_gbp_numeric
        stock = Normalize-CellText $row.stock
        delivery_delay = Normalize-CellText $row.delivery_delay
        url = Normalize-CellText $row.url
    }
}

$cleanRows | Select-Object $columnOrder | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$excel = $null
$workbook = $null
$worksheet = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if (Test-Path $OutputXlsx) {
        Remove-Item $OutputXlsx -Force
    }

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = 'Products'

    for ($columnIndex = 0; $columnIndex -lt $columnOrder.Count; $columnIndex++) {
        $worksheet.Cells.Item(1, $columnIndex + 1).Value2 = $columnOrder[$columnIndex]
    }

    $rowIndex = 2
    foreach ($row in $cleanRows) {
        for ($columnIndex = 0; $columnIndex -lt $columnOrder.Count; $columnIndex++) {
            $propertyName = $columnOrder[$columnIndex]
            $worksheet.Cells.Item($rowIndex, $columnIndex + 1).Value2 = $row.$propertyName
        }
        $rowIndex++
    }

    $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $columnOrder.Count))
    $headerRange.Font.Bold = $true
    $worksheet.Rows.Item(1).WrapText = $false
    $worksheet.UsedRange.WrapText = $false
    $worksheet.Columns.Item(5).NumberFormat = '@'

    for ($columnIndex = 1; $columnIndex -le $columnOrder.Count; $columnIndex++) {
        Set-ColumnWidthCm -worksheet $worksheet -excel $excel -columnIndex $columnIndex -centimeters 5
    }

    $workbook.SaveAs([System.IO.Path]::GetFullPath($OutputXlsx), 51)
}
finally {
    if ($workbook) {
        $workbook.Close($true)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    if ($worksheet) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    }

    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host "CSV repare ecrit dans $OutputCsv"
Write-Host "XLSX formate ecrit dans $OutputXlsx"