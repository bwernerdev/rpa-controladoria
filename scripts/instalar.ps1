# Instala ou atualiza o Robô SRP para o usuário atual (sem administrador) e cria os atalhos.
# Vai dentro do zip do release; rode pelo instalar.cmd (duplo clique) na pasta extraída.
# Preserva o que já existe no destino: config.json, aprendizado.json, progresso.json, logs\ e planilhas\.
param(
    [string]$Destino = "$env:LOCALAPPDATA\Programs\RPA-SRP",
    [string[]]$PastasAtalho = @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs"))
)
$ErrorActionPreference = "Stop"
$origem = $PSScriptRoot

if (-not (Test-Path "$origem\RPA-SRP.exe")) { throw "RPA-SRP.exe não encontrado em $origem. Extraia o zip inteiro antes de instalar." }
New-Item -ItemType Directory -Force $Destino, "$Destino\planilhas" | Out-Null
try {
    Copy-Item "$origem\RPA-SRP.exe" $Destino -Force
} catch {
    throw "Não consegui substituir o RPA-SRP.exe: feche o Robô SRP e rode o instalador de novo."
}
if (Test-Path "$Destino\tesseract") { Remove-Item "$Destino\tesseract" -Recurse -Force }
Copy-Item "$origem\tesseract" "$Destino\tesseract" -Recurse
if (Test-Path "$origem\planilhas\BASE PADRAO.xlsx") {
    Copy-Item "$origem\planilhas\BASE PADRAO.xlsx" "$Destino\planilhas" -Force
}

$shell = New-Object -ComObject WScript.Shell
foreach ($pasta in $PastasAtalho) {
    if (-not $pasta) { continue }
    New-Item -ItemType Directory -Force $pasta | Out-Null
    $atalho = $shell.CreateShortcut((Join-Path $pasta "Robô SRP.lnk"))
    $atalho.TargetPath = "$Destino\RPA-SRP.exe"
    $atalho.WorkingDirectory = $Destino
    $atalho.IconLocation = "$Destino\RPA-SRP.exe,0"
    $atalho.Description = "Preenche o SRP a partir de uma planilha"
    $atalho.Save()
}
Write-Host "Robô SRP instalado em $Destino"
Write-Host "Atalho criado na área de trabalho e no menu Iniciar."
