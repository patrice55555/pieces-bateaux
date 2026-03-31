param(
    [string]$InputXlsx = ".\princess_parts_products.xlsx",
    [string]$DbPath = ".\princess_parts_analytics.sqlite.db",
    [string]$WorksheetName,
    [string]$GeneratorScript = ".\scripts\generate-sqlite-analytics-sql.ps1",
    [switch]$KeepTempFiles
)

$ErrorActionPreference = 'Stop'

function Get-SqliteExecutable {
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
    $sqliteCommand = Get-Command sqlite3 -ErrorAction SilentlyContinue
    if (-not $sqliteCommand) {
        throw 'sqlite3 est introuvable dans le PATH.'
    }

    return $sqliteCommand.Source
}

function Convert-ToCsvField {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        $Value = ''
    }

    $escaped = $Value.Replace('"', '""')
    return '"' + $escaped + '"'
}

function Export-WorksheetToCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputCsv,

        [string]$SheetName
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null
    $writer = $null
    $numericHeaders = @('price_gbp_numeric', 'price_euro_numeric', 'price_dollar_numeric', 'stock')

    function Get-WorksheetCellText {
        param(
            [Parameter(Mandatory = $true)]$Sheet,
            [int]$Row,
            [int]$Column,
            [string]$HeaderName
        )

        $cell = $null
        try {
            $cell = $Sheet.Cells.Item($Row, $Column)
            $rawValue = $cell.Value2

            if ($Row -eq 1) {
                return [string]$cell.Text
            }

            if ($HeaderName -eq 'date') {
                if ($rawValue -is [double] -or $rawValue -is [int]) {
                    return ([DateTime]::FromOADate([double]$rawValue)).ToString('yyyy-MM-dd')
                }

                return [string]$cell.Text
            }

            if ($numericHeaders -contains $HeaderName) {
                if ($null -eq $rawValue -or [string]::IsNullOrWhiteSpace([string]$rawValue)) {
                    return ''
                }

                if ($HeaderName -eq 'stock') {
                    return ([int][double]$rawValue).ToString([System.Globalization.CultureInfo]::InvariantCulture)
                }

                return ([double]$rawValue).ToString('0.00', [System.Globalization.CultureInfo]::InvariantCulture)
            }

            return [string]$cell.Text
        }
        finally {
            if ($cell) {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($cell) | Out-Null
            }
        }
    }

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
        $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, $null, $true)

        if ([string]::IsNullOrWhiteSpace($SheetName)) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        else {
            $worksheet = $workbook.Worksheets.Item($SheetName)
        }

        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count
        $columnCount = $usedRange.Columns.Count
        $headers = @()

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($OutputCsv, $false, $utf8NoBom)

        for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $headers += (Get-WorksheetCellText -Sheet $worksheet -Row 1 -Column $columnIndex -HeaderName '')
        }

        $writer.WriteLine((($headers | ForEach-Object { Convert-ToCsvField $_ }) -join ','))

        for ($rowIndex = 2; $rowIndex -le $rowCount; $rowIndex++) {
            $fields = @()
            for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
                $headerName = $headers[$columnIndex - 1]
                $fields += Convert-ToCsvField (Get-WorksheetCellText -Sheet $worksheet -Row $rowIndex -Column $columnIndex -HeaderName $headerName)
            }
            $writer.WriteLine(($fields -join ','))
        }
    }
    finally {
        if ($writer) {
            $writer.Dispose()
        }
        if ($usedRange) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($usedRange) | Out-Null
        }
        if ($workbook) {
            $workbook.Close($false)
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
}

if (-not (Test-Path -LiteralPath $InputXlsx)) {
    throw "Fichier Excel introuvable: $InputXlsx"
}

if (-not (Test-Path -LiteralPath $GeneratorScript)) {
    throw "Script generateur introuvable: $GeneratorScript"
}

$dbDirectory = Split-Path -Parent $DbPath
if (-not [string]::IsNullOrWhiteSpace($dbDirectory) -and -not (Test-Path -LiteralPath $dbDirectory)) {
    New-Item -ItemType Directory -Path $dbDirectory -Force | Out-Null
}

$tempBase = Join-Path ([System.IO.Path]::GetTempPath()) ("princess_parts_" + [guid]::NewGuid().ToString('N'))
$tempCsv = "$tempBase.csv"
$tempSql = "$tempBase.sql"

try {
    Export-WorksheetToCsv -WorkbookPath $InputXlsx -OutputCsv $tempCsv -SheetName $WorksheetName

    & $GeneratorScript -InputCsv $tempCsv -OutputSql $tempSql -SourceFileLabel (Split-Path -Leaf $InputXlsx)

    $sqliteExe = Get-SqliteExecutable
    & $sqliteExe $DbPath ".read $tempSql"
    & $sqliteExe $DbPath "DELETE FROM import_runs WHERE import_run_id NOT IN (SELECT DISTINCT import_run_id FROM princess_parts_history WHERE import_run_id IS NOT NULL);"

    $historyCount = (& $sqliteExe $DbPath "SELECT COUNT(*) FROM princess_parts_history;") | Select-Object -Last 1
    $runCount = (& $sqliteExe $DbPath "SELECT COUNT(*) FROM import_runs;") | Select-Object -Last 1
    $maxSnapshotDate = (& $sqliteExe $DbPath "SELECT MAX(snapshot_date) FROM princess_parts_history;") | Select-Object -Last 1

    Write-Host "Base SQLite ecrite dans $DbPath"
    Write-Host "Historique total: $historyCount ligne(s)"
    Write-Host "Imports traces: $runCount"
    Write-Host "Dernier snapshot: $maxSnapshotDate"
}
finally {
    if (-not $KeepTempFiles) {
        if (Test-Path -LiteralPath $tempCsv) {
            Remove-Item -LiteralPath $tempCsv -Force
        }
        if (Test-Path -LiteralPath $tempSql) {
            Remove-Item -LiteralPath $tempSql -Force
        }
    }
}