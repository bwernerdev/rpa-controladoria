# Instala ou atualiza o Robô SRP para o usuário atual (sem administrador) e cria os atalhos.
# Vai dentro do zip do release; rode pelo instalar.cmd (duplo clique) na pasta extraída.
# Preserva o que já existe no destino: config.json, aprendizado.json, progresso.json, logs\ e planilhas\.
# A autoatualização do app chama este script com -SemAtalhos -Esperar 60 -Abrir -Log <arquivo>.
param(
    [string]$Destino = "$env:LOCALAPPDATA\Programs\RPA-SRP",
    [string[]]$PastasAtalho = @([Environment]::GetFolderPath("Desktop"), [Environment]::GetFolderPath("Programs")),
    [switch]$SemAtalhos,
    [int]$Esperar = 0,   # segundos tentando trocar o exe enquanto o app que pediu a atualização fecha
    [switch]$Abrir,      # abre o app no fim (inclusive se a atualização falhar, para não deixar o usuário sem ele)
    [string]$Log = ""
)
$ErrorActionPreference = "Stop"
$origem = $PSScriptRoot
if ($Log) { Start-Transcript -Path $Log -Append | Out-Null }
try {
    if (-not (Test-Path "$origem\RPA-SRP.exe")) { throw "RPA-SRP.exe não encontrado em $origem. Extraia o zip inteiro antes de instalar." }
    New-Item -ItemType Directory -Force $Destino, "$Destino\planilhas" | Out-Null
    $prazo = (Get-Date).AddSeconds($Esperar)
    while ($true) {
        try {
            Copy-Item "$origem\RPA-SRP.exe" $Destino -Force
            break
        } catch {
            if ((Get-Date) -ge $prazo) { throw "Não consegui substituir o RPA-SRP.exe: feche o Robô SRP e rode o instalador de novo." }
            Start-Sleep -Milliseconds 500
        }
    }
    if (Test-Path "$Destino\tesseract") { Remove-Item "$Destino\tesseract" -Recurse -Force }
    Copy-Item "$origem\tesseract" "$Destino\tesseract" -Recurse
    if (Test-Path "$origem\planilhas\BASE PADRAO.xlsx") {
        Copy-Item "$origem\planilhas\BASE PADRAO.xlsx" "$Destino\planilhas" -Force
    }
    # O zip baixado vem com a marca "da internet", que passa para o que é extraído e copiado; com ela o
    # SmartScreen mostra a tela azul ao abrir o exe instalado. O usuário já escolheu instalar: tira a marca
    # só do que este script instalou (o mesmo que "Desbloquear" nas Propriedades do arquivo).
    try {
        @(Get-Item "$Destino\RPA-SRP.exe") + @(Get-ChildItem "$Destino\tesseract" -Recurse -File) +
            @(Get-ChildItem "$Destino\planilhas\BASE PADRAO.xlsx" -ErrorAction SilentlyContinue) | Unblock-File
    } catch {
        Write-Host "Aviso: não consegui desbloquear os arquivos ($($_.Exception.Message)). Se o Windows avisar ao abrir, clique em 'Mais informações' e 'Executar assim mesmo'."
    }

    if (-not $SemAtalhos) {
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
        Write-Host "Atalho criado na área de trabalho e no menu Iniciar."
    }
    Write-Host "Robô SRP instalado em $Destino"
} finally {
    if ($Abrir -and (Test-Path "$Destino\RPA-SRP.exe")) {
        # Herdado do app antigo (exe do PyInstaller): sem limpar, o exe novo procura o python314.dll na pasta _MEI do antigo, que ja foi apagada.
        @(Get-ChildItem Env: | Where-Object { $_.Name -like "_PYI_*" -or $_.Name -like "_MEIPASS*" }) | ForEach-Object { Remove-Item "Env:$($_.Name)" }
        $env:PYINSTALLER_RESET_ENVIRONMENT = "1"
        Start-Process "$Destino\RPA-SRP.exe" -WorkingDirectory $Destino
    }
    if ($Log) { Stop-Transcript | Out-Null }
}
