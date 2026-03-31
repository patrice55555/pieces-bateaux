param(
    [string]$InputCsv = ".\princess_parts_products.csv",
    [string]$OutputCsv = ".\princess_parts_products.csv",
    [string]$OutputXlsx = ".\princess_parts_products.xlsx",
    [string]$OutputSql = ".\princess_parts_products.sql"
)

$InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
$SnapshotDate = (Get-Date).ToString('yyyy-MM-dd', $InvariantCulture)

$columnOrder = @(
    'date',
    'category',
    'Category (FR)',
    'name',
    'Name (FR)',
    'description',
    'Description (FR)',
    'specifications',
    'Specifications (FR)',
    'model_number',
    'price_gbp_numeric',
    'price_euro_numeric',
    'price_dollar_numeric',
    'stock',
    'delivery_delay',
    'Delivery delay (FR)',
    'url'
)

$categoryMap = @{
    'Anodes' = 'Anodes'
    'Boat maintenance' = 'Entretien du bateau'
    'Branded Clothing' = 'Vetements de marque'
    'Clearance items' = 'Articles en destockage'
    'Electrical' = 'Electricite'
    'Entertainment' = 'Divertissement embarque'
    'External Hardware' = 'Accastillage exterieur'
    'External Protection' = 'Protection exterieure'
    'Generator Service parts' = 'Pieces d''entretien de generateur'
    'Gifts and Merchandise' = 'Cadeaux et articles de marque'
    'Halyard' = 'Halyard'
    'Interior Hardware' = 'Quincaillerie interieure'
    'Lighting' = 'Eclairage'
    'Plumbing & Water Systems' = 'Plomberie et circuits d''eau'
    'Propellers' = 'Helices'
    'Regular Servicing' = 'Entretien courant'
    'Tender mounting' = 'Fixations pour annexe'
    'Yacht Accessories' = 'Accessoires pour yacht'
}

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

function Convert-ToFrenchDecimalText([string]$Value) {
    return [regex]::Replace($Value, '(?<int>\d+)\.(?<dec>\d+)', '$1,$2')
}

function Convert-ToSqlLiteral([string]$Value) {
    if ($null -eq $Value) {
        return 'NULL'
    }

    $normalized = Normalize-CellText $Value
    return "'" + $normalized.Replace("'", "''") + "'"
}

function Export-SqlSnapshot([object[]]$Rows, [string]$Path) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('-- Princess Parts daily snapshot SQL') | Out-Null
    $lines.Add('-- Generated on ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $InvariantCulture)) | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('CREATE TABLE IF NOT EXISTS princess_parts_history (') | Out-Null
    $lines.Add('    snapshot_date DATE NOT NULL,') | Out-Null
    $lines.Add('    category TEXT,') | Out-Null
    $lines.Add('    category_fr TEXT,') | Out-Null
    $lines.Add('    name TEXT,') | Out-Null
    $lines.Add('    name_fr TEXT,') | Out-Null
    $lines.Add('    description TEXT,') | Out-Null
    $lines.Add('    description_fr TEXT,') | Out-Null
    $lines.Add('    specifications TEXT,') | Out-Null
    $lines.Add('    specifications_fr TEXT,') | Out-Null
    $lines.Add('    model_number TEXT,') | Out-Null
    $lines.Add('    price_gbp_numeric NUMERIC(12,2),') | Out-Null
    $lines.Add('    price_euro_numeric NUMERIC(12,2),') | Out-Null
    $lines.Add('    price_dollar_numeric NUMERIC(12,2),') | Out-Null
    $lines.Add('    stock INTEGER,') | Out-Null
    $lines.Add('    delivery_delay TEXT,') | Out-Null
    $lines.Add('    delivery_delay_fr TEXT,') | Out-Null
    $lines.Add('    url TEXT NOT NULL,') | Out-Null
    $lines.Add('    PRIMARY KEY (snapshot_date, url)') | Out-Null
    $lines.Add(');') | Out-Null
    $lines.Add('') | Out-Null
    $lines.Add('CREATE INDEX IF NOT EXISTS idx_princess_parts_history_url ON princess_parts_history (url);') | Out-Null
    $lines.Add('CREATE INDEX IF NOT EXISTS idx_princess_parts_history_model_number ON princess_parts_history (model_number);') | Out-Null
    $lines.Add('') | Out-Null

    foreach ($row in $Rows) {
        $lines.Add(
            'INSERT INTO princess_parts_history (' +
            'snapshot_date, category, category_fr, name, name_fr, description, description_fr, specifications, specifications_fr, model_number, ' +
            'price_gbp_numeric, price_euro_numeric, price_dollar_numeric, stock, delivery_delay, delivery_delay_fr, url' +
            ') VALUES (' +
            (Convert-ToSqlLiteral $row.date) + ', ' +
            (Convert-ToSqlLiteral $row.category) + ', ' +
            (Convert-ToSqlLiteral $row.'Category (FR)') + ', ' +
            (Convert-ToSqlLiteral $row.name) + ', ' +
            (Convert-ToSqlLiteral $row.'Name (FR)') + ', ' +
            (Convert-ToSqlLiteral $row.description) + ', ' +
            (Convert-ToSqlLiteral $row.'Description (FR)') + ', ' +
            (Convert-ToSqlLiteral $row.specifications) + ', ' +
            (Convert-ToSqlLiteral $row.'Specifications (FR)') + ', ' +
            (Convert-ToSqlLiteral $row.model_number) + ', ' +
            (Convert-ToSqlLiteral $row.price_gbp_numeric) + ', ' +
            (Convert-ToSqlLiteral $row.price_euro_numeric) + ', ' +
            (Convert-ToSqlLiteral $row.price_dollar_numeric) + ', ' +
            (Convert-ToSqlLiteral $row.stock) + ', ' +
            (Convert-ToSqlLiteral $row.delivery_delay) + ', ' +
            (Convert-ToSqlLiteral $row.'Delivery delay (FR)') + ', ' +
            (Convert-ToSqlLiteral $row.url) +
            ');'
        ) | Out-Null
    }

    $lines | Set-Content -Path $Path -Encoding UTF8
}

function Get-LatestExchangeRates() {
    $fallback = @{
        EUR = 1.150601
        USD = 1.319612
        Timestamp = 'Tue, 31 Mar 2026 00:02:31 +0000'
        Source = 'fallback'
    }

    try {
        $response = Invoke-RestMethod -Uri 'https://open.er-api.com/v6/latest/GBP' -TimeoutSec 60
        if ($response.result -eq 'success' -and $response.rates.EUR -and $response.rates.USD) {
            return @{
                EUR = [decimal]::Parse(([string]$response.rates.EUR), $InvariantCulture)
                USD = [decimal]::Parse(([string]$response.rates.USD), $InvariantCulture)
                Timestamp = [string]$response.time_last_update_utc
                Source = 'open.er-api.com'
            }
        }
    }
    catch {
    }

    return @{
        EUR = [decimal]$fallback.EUR
        USD = [decimal]$fallback.USD
        Timestamp = [string]$fallback.Timestamp
        Source = [string]$fallback.Source
    }
}

function Convert-Price([string]$GbpValue, [decimal]$Rate) {
    $normalized = Normalize-CellText $GbpValue
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return '0.00'
    }

    $amount = [decimal]::Parse($normalized, $InvariantCulture)
    return (($amount * $Rate).ToString('0.00', $InvariantCulture))
}

function Get-CategoryFr([string]$Category) {
    $normalized = Normalize-CellText $Category
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $parts = $normalized -split '\s*\|\s*'
    $translated = foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($categoryMap.ContainsKey($trimmed)) {
            $categoryMap[$trimmed]
        }
        else {
            $trimmed
        }
    }

    return ($translated -join ' | ')
}

function Get-SpecificationsFr([string]$Specifications) {
    $normalized = Normalize-CellText $Specifications
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $result = $normalized
    $result = $result -replace '^Dimensions\s*-\s*', 'Dimensions : '
    $result = $result -replace '\s+Weight\s*-\s*', ' | Poids : '
    $result = Convert-ToFrenchDecimalText $result
    return $result
}

$rows = Import-Csv $InputCsv
$rates = Get-LatestExchangeRates

$finalRows = foreach ($row in $rows) {
    $category = Normalize-CellText $row.category
    $specifications = Normalize-CellText $row.specifications
    $priceGbpNumeric = Normalize-CellText $row.price_gbp_numeric

    [pscustomobject]@{
        date = $SnapshotDate
        category = $category
        'Category (FR)' = Get-CategoryFr $category
        name = Normalize-CellText $row.name
        'Name (FR)' = Normalize-CellText $row.'Name (FR)'
        description = Normalize-CellText $row.description
        'Description (FR)' = Normalize-CellText $row.'Description (FR)'
        specifications = $specifications
        'Specifications (FR)' = Get-SpecificationsFr $specifications
        model_number = Normalize-CellText $row.model_number
        price_gbp_numeric = $priceGbpNumeric
        price_euro_numeric = Convert-Price -GbpValue $priceGbpNumeric -Rate $rates.EUR
        price_dollar_numeric = Convert-Price -GbpValue $priceGbpNumeric -Rate $rates.USD
        stock = Normalize-CellText $row.stock
        delivery_delay = Normalize-CellText $row.delivery_delay
        'Delivery delay (FR)' = Normalize-CellText $row.'Delivery delay (FR)'
        url = Normalize-CellText $row.url
    }
}

$finalRows | Select-Object $columnOrder | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
Export-SqlSnapshot -Rows $finalRows -Path $OutputSql

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
    foreach ($row in $finalRows) {
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
    $worksheet.Columns.Item(10).NumberFormat = '@'
    $worksheet.Columns.Item(11).NumberFormat = '0.00'
    $worksheet.Columns.Item(12).NumberFormat = '0.00'
    $worksheet.Columns.Item(13).NumberFormat = '0.00'

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

Write-Host "CSV metier ecrit dans $OutputCsv"
Write-Host "XLSX metier ecrit dans $OutputXlsx"
Write-Host "SQL metier ecrit dans $OutputSql"
Write-Host ("Taux utilises GBP->EUR=" + $rates.EUR.ToString('0.000000', $InvariantCulture) + ", GBP->USD=" + $rates.USD.ToString('0.000000', $InvariantCulture) + ", date=" + $rates.Timestamp + ", source=" + $rates.Source)