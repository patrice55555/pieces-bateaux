param(
    [string]$OutputFile = "princess_parts_products.csv"
)

$ProgressPreference = 'SilentlyContinue'
$InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
$ColumnOrder = @(
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

function Decode-Html([string]$Value) {
    return [System.Net.WebUtility]::HtmlDecode($Value)
}

function Normalize-Text([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    $decoded = Decode-Html $Value
    return (($decoded -replace '<[^>]+>', ' ') -replace '\s+', ' ').Trim()
}

function Get-FirstMatchValue([string]$Html, [string]$Pattern) {
    $match = [regex]::Match($Html, $Pattern, [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
    if ($match.Success) {
        return Normalize-Text $match.Groups['v'].Value
    }

    return ''
}

function Get-AllMatchValues([string]$Html, [string]$Pattern) {
    $matches = [regex]::Matches($Html, $Pattern, [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
    $values = New-Object System.Collections.Generic.List[string]

    foreach ($match in $matches) {
        $value = Normalize-Text $match.Groups['v'].Value
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            [void]$values.Add($value)
        }
    }

    return $values
}

function Get-LinkTexts([string]$Html) {
    $matches = [regex]::Matches($Html, '<a[^>]*>(?<v>.*?)</a>', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
    $values = New-Object System.Collections.Generic.List[string]

    foreach ($match in $matches) {
        $value = Normalize-Text $match.Groups['v'].Value
        if (-not [string]::IsNullOrWhiteSpace($value) -and -not $values.Contains($value)) {
            [void]$values.Add($value)
        }
    }

    return $values
}

function Join-Unique([System.Collections.Generic.List[string]]$Values, [string]$Separator) {
    if ($null -eq $Values -or $Values.Count -eq 0) {
        return ''
    }

    return ($Values | Select-Object -Unique) -join $Separator
}

function Normalize-Stock([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return '0'
    }

    $match = [regex]::Match($Value, '(?<qty>\d+)')
    if ($match.Success) {
        return $match.Groups['qty'].Value
    }

    return '0'
}

function Normalize-Price([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return '0.00'
    }

    $clean = $Value.Replace(',', '')
    $match = [regex]::Match($clean, '(?<price>\d+(?:\.\d+)?)')
    if ($match.Success) {
        $decimalPrice = [decimal]::Parse($match.Groups['price'].Value, $InvariantCulture)
        return $decimalPrice.ToString('0.00', $InvariantCulture)
    }

    return '0.00'
}

$siteMapResponse = Invoke-WebRequest -Uri 'https://parts.princess.co.uk/product-sitemap.xml' -UseBasicParsing -TimeoutSec 60
$productUrls = [regex]::Matches($siteMapResponse.Content, '<loc>(?<v>.*?)</loc>', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline') |
    ForEach-Object { $_.Groups['v'].Value.Trim() } |
    Where-Object { $_ -like 'https://parts.princess.co.uk/product/*' }

$results = New-Object System.Collections.Generic.List[object]
$total = $productUrls.Count
$index = 0

foreach ($productUrl in $productUrls) {
    $index += 1

    try {
        $html = (Invoke-WebRequest -Uri $productUrl -UseBasicParsing -TimeoutSec 60).Content

        $name = Get-FirstMatchValue $html '<h1 class="product_title entry-title">(?<v>.*?)</h1>'
        $price = Get-FirstMatchValue $html '<p class="price">.*?woocommerce-Price-currencySymbol">&pound;</span>(?<v>[0-9.,]+)</bdi>'
        $priceNumeric = Normalize-Price $price
        $stock = Normalize-Stock (Get-FirstMatchValue $html '<p class="stock[^"]*">(?<v>.*?)</p>')
        $modelNumber = Get-FirstMatchValue $html '<span class="sku_wrapper">\s*Model Number:\s*<span class="sku">(?<v>.*?)</span>'

        $categoriesHtmlMatch = [regex]::Match($html, '<span class="posted_in">\s*Categor(?:y|ies):\s*(?<v>.*?)</span>', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
        $categories = New-Object System.Collections.Generic.List[string]
        if ($categoriesHtmlMatch.Success) {
            $categories = Get-LinkTexts $categoriesHtmlMatch.Groups['v'].Value
        }

        $descriptionParts = New-Object System.Collections.Generic.List[string]
        $descriptionBlockMatch = [regex]::Match($html, '<div class="woocommerce-product-details__short-description">(?<v>.*?)</div>\s*<div class="product_meta">', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
        if ($descriptionBlockMatch.Success) {
            $paragraphs = Get-AllMatchValues $descriptionBlockMatch.Groups['v'].Value '<p[^>]*>(?<v>.*?)</p>'
            foreach ($paragraph in $paragraphs) {
                [void]$descriptionParts.Add($paragraph)
            }
        }

        $specifications = ''
        $deliveryDelay = ''
        $accordionMatch = [regex]::Match($html, '<ul id="accordion">(?<v>.*?)</ul>', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
        if ($accordionMatch.Success) {
            $sectionMatches = [regex]::Matches($accordionMatch.Groups['v'].Value, '<li>\s*<h4>(?<title>.*?)</h4>\s*<div class="content">(?<content>.*?)</div>\s*</li>', [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline')
            foreach ($section in $sectionMatches) {
                $title = Normalize-Text $section.Groups['title'].Value
                $content = Normalize-Text $section.Groups['content'].Value

                if ($title -eq 'Specification') {
                    $specifications = $content
                    continue
                }

                if ($title -eq 'Delivery') {
                    $deliveryDelay = $content
                }
            }
        }

        $results.Add([pscustomobject]@{
            category = (Join-Unique $categories ' | ')
            name = $name
            description = (Join-Unique $descriptionParts ' | ')
            specifications = $specifications
            model_number = $modelNumber
            price_gbp_numeric = $priceNumeric
            stock = $stock
            delivery_delay = $deliveryDelay
            url = $productUrl
        }) | Out-Null
    }
    catch {
        $results.Add([pscustomobject]@{
            category = ''
            name = ''
            description = ''
            specifications = ''
            model_number = ''
            price_gbp_numeric = '0.00'
            stock = ''
            delivery_delay = ''
            url = $productUrl
        }) | Out-Null
    }

    if (($index % 25) -eq 0 -or $index -eq $total) {
        Write-Host "Traite $index / $total"
    }

    Start-Sleep -Milliseconds 150
}

$results | Select-Object $ColumnOrder | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Host "Produits: $($results.Count)"
Write-Host "CSV ecrit dans $OutputFile"