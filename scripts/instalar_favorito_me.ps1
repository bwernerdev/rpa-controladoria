# Instala o favorito "ME Friboi V2" na barra de favoritos do Chrome ou Edge.
# Feche o navegador antes (o script se recusa se ele estiver aberto; nao encerra processos).
# Uso: powershell -NoProfile -ExecutionPolicy Bypass -File .\scripts\instalar_favorito_me.ps1 [-Navegador chrome|edge] [-Perfil Default]
param(
    [ValidateSet('chrome', 'edge')][string]$Navegador = 'chrome',
    [string]$Perfil = 'Default',
    [string]$Nome = 'ME Friboi V2',
    [string]$Arquivo   # so para teste: caminho de um arquivo Bookmarks
)
$ErrorActionPreference = 'Stop'

$url = (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'favorito_me_friboi_v2.txt') -Raw -Encoding UTF8).Trim()
if (-not $url.StartsWith('javascript:')) { throw 'favorito_me_friboi_v2.txt invalido.' }

if (-not $Arquivo) {
    $proc = if ($Navegador -eq 'chrome') { 'chrome' } else { 'msedge' }
    $base = if ($Navegador -eq 'chrome') { "$env:LOCALAPPDATA\Google\Chrome\User Data" } else { "$env:LOCALAPPDATA\Microsoft\Edge\User Data" }
    if (Get-Process -Name $proc -ErrorAction SilentlyContinue) {
        throw "Feche o $Navegador por completo (inclusive na bandeja) e rode de novo."
    }
    $Arquivo = Join-Path $base "$Perfil\Bookmarks"
}
if (-not (Test-Path -LiteralPath $Arquivo)) { throw "Nao achei $Arquivo (perfil errado? use -Perfil 'Profile 1')." }

Copy-Item -LiteralPath $Arquivo -Destination "$Arquivo.bak-me" -Force
$json = Get-Content -LiteralPath $Arquivo -Raw -Encoding UTF8 | ConvertFrom-Json

function Max-Id($no) {
    $m = [int]$no.id
    if ($no.children) { foreach ($c in $no.children) { $m = [Math]::Max($m, (Max-Id $c)) } }
    $m
}
$maxId = 0
foreach ($raiz in $json.roots.PSObject.Properties) {
    if ($raiz.Value -and $raiz.Value.id) { $maxId = [Math]::Max($maxId, (Max-Id $raiz.Value)) }
}

$barra = $json.roots.bookmark_bar
$existente = @($barra.children | Where-Object { $_.name -eq $Nome })
if ($existente.Count) {
    foreach ($e in $existente) { $e.url = $url }
    $acao = 'atualizado'
} else {
    $novo = [ordered]@{
        date_added = [string]([DateTime]::UtcNow.ToFileTimeUtc() / 10)
        guid       = [guid]::NewGuid().ToString()
        id         = [string]($maxId + 1)
        name       = $Nome
        type       = 'url'
        url        = $url
    }
    $barra.children = @($barra.children) + [pscustomobject]$novo
    $acao = 'adicionado'
}
# o Chrome recalcula o checksum ao abrir; um valor velho nao deve ficar
if ($json.PSObject.Properties['checksum']) { $json.PSObject.Properties.Remove('checksum') }

[IO.File]::WriteAllText($Arquivo, ($json | ConvertTo-Json -Depth 100 -Compress), (New-Object Text.UTF8Encoding($false)))
Write-Host "Favorito '$Nome' $acao em $Arquivo (copia: $Arquivo.bak-me). Abra o navegador e mostre a barra de favoritos (Ctrl+Shift+B)."
