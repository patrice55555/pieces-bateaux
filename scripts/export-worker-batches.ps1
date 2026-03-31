param(
    [string]$BaseUrl,
    [string]$OutputFile = "princess_parts_products.csv",
    [int]$BatchSize = 20
)

if (-not $BaseUrl) {
    throw "Passez l'URL du Worker via -BaseUrl, par exemple https://princess-parts-scraper.<votre-sous-domaine>.workers.dev"
}

$manifest = Invoke-RestMethod -Uri "$BaseUrl/manifest"
$total = [int]$manifest.total

if ($total -le 0) {
    throw "Le Worker n'a retourne aucun produit dans le manifest."
}

$headerWritten = $false
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile -Force
}

for ($offset = 0; $offset -lt $total; $offset += $BatchSize) {
    $uri = "$BaseUrl/scrape.csv?offset=$offset&limit=$BatchSize"
    $response = Invoke-WebRequest -Uri $uri
    $lines = ($response.Content -split "`r?`n") | Where-Object { $_ -ne "" }

    if (-not $headerWritten) {
        $lines | Set-Content -Path $OutputFile -Encoding utf8
        $headerWritten = $true
        continue
    }

    $lines | Select-Object -Skip 1 | Add-Content -Path $OutputFile -Encoding utf8
}

Write-Host "CSV consolide ecrit dans $OutputFile"