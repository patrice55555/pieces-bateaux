param(
    [string]$InputCsv = ".\princess_parts_products.csv",
    [string]$OutputCsv = ".\princess_parts_products.csv",
    [string]$OutputXlsx = ".\princess_parts_products.xlsx",
    [string]$CacheFile = ".\.translation_cache_fr.json"
)

$ProgressPreference = 'SilentlyContinue'

$columnOrder = @(
    'category',
    'Category (FR)',
    'name',
    'Name (FR)',
    'description',
    'Description (FR)',
    'specifications',
    'model_number',
    'price_gbp_numeric',
    'stock',
    'delivery_delay',
    'Delivery delay (FR)',
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

function Load-TranslationCache([string]$Path) {
    $cache = @{}

    if (-not (Test-Path $Path)) {
        return $cache
    }

    $raw = Get-Content $Path -Raw -Encoding UTF8
    if ([string]::IsNullOrWhiteSpace($raw)) {
        return $cache
    }

    $data = $raw | ConvertFrom-Json -AsHashtable
    if ($data) {
        foreach ($key in $data.Keys) {
            $cache[$key] = [string]$data[$key]
        }
    }

    return $cache
}

function Save-TranslationCache([hashtable]$Cache, [string]$Path) {
    $Cache | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8
}

function Translate-Text([string]$Value, [hashtable]$Cache) {
    $normalized = Normalize-CellText $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    if ($Cache.ContainsKey($normalized)) {
        return $Cache[$normalized]
    }

    $url = 'https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=fr&dt=t&q=' + [uri]::EscapeDataString($normalized)
    $response = Invoke-RestMethod -Uri $url -TimeoutSec 60
    $translated = (($response[0] | ForEach-Object { $_[0] }) -join '')
    $translated = Normalize-CellText $translated

    $Cache[$normalized] = $translated
    return $translated
}

$rows = Import-Csv $InputCsv
$cache = Load-TranslationCache $CacheFile
$uniqueTexts = New-Object System.Collections.Generic.HashSet[string]

foreach ($row in $rows) {
    [void]$uniqueTexts.Add((Normalize-CellText $row.category))
    [void]$uniqueTexts.Add((Normalize-CellText $row.name))
    [void]$uniqueTexts.Add((Normalize-CellText $row.description))
    [void]$uniqueTexts.Add((Normalize-CellText $row.delivery_delay))
}

$allTexts = $uniqueTexts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$translatedCount = 0

foreach ($text in $allTexts) {
    if (-not $cache.ContainsKey($text)) {
        $translated = Translate-Text -Value $text -Cache $cache
        $translatedCount += 1

        if (($translatedCount % 25) -eq 0) {
            Save-TranslationCache -Cache $cache -Path $CacheFile
            Write-Host "Traductions effectuees: $translatedCount"
        }

        Start-Sleep -Milliseconds 120
    }
}

Save-TranslationCache -Cache $cache -Path $CacheFile

$translatedRows = foreach ($row in $rows) {
    $category = Normalize-CellText $row.category
    $name = Normalize-CellText $row.name
    $description = Normalize-CellText $row.description
    $deliveryDelay = Normalize-CellText $row.delivery_delay

    [pscustomobject]@{
        category = $category
        'Category (FR)' = $cache[$category]
        name = $name
        'Name (FR)' = $cache[$name]
        description = $description
        'Description (FR)' = $cache[$description]
        specifications = Normalize-CellText $row.specifications
        model_number = Normalize-CellText $row.model_number
        price_gbp_numeric = Normalize-CellText $row.price_gbp_numeric
        stock = Normalize-CellText $row.stock
        delivery_delay = $deliveryDelay
        'Delivery delay (FR)' = $cache[$deliveryDelay]
        url = Normalize-CellText $row.url
    }
}

$translatedRows | Select-Object $columnOrder | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

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
    foreach ($row in $translatedRows) {
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
    $worksheet.Columns.Item(8).NumberFormat = '@'

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

Write-Host "CSV traduit ecrit dans $OutputCsv"
Write-Host "XLSX traduit ecrit dans $OutputXlsx"