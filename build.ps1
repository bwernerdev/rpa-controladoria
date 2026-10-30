# Gera build\dist\RPA-SRP.exe (arquivo unico, com tela de abertura, icone e versao). Rode na pasta RPA.
# O exe acha o Tesseract numa pasta tesseract ao lado dele (como no zip/instalacao); a raiz do projeto fica sem exe.
# Abertura medida (30/10/2026): ~4,6 s, quase tudo antes do Python comecar (verificacao do Windows/antivirus);
# o formato de pasta (--onedir) deu o mesmo tempo, por isso ficou o arquivo unico.
# Depois do exe, monta release\RPA-SRP-vX.Y.Z.zip (exe + tesseract\ + BASE PADRAO.xlsx + instalar.cmd) para o GitHub Releases;
# o exe e o zip nao vao ao git.
$raiz = $PWD.Path
& "$raiz\.venv\Scripts\python.exe" src\icone.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_versao.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_splash.py  # a tela de abertura mostra a versao: refeita a cada build
& "$raiz\.venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --splash "$raiz\build\splash.png" --icon "$raiz\assets\icone.ico" --version-file "$raiz\build\version.txt" --distpath "$raiz\build\dist" --workpath build --specpath build src\gui.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou" }
$ErrorActionPreference = "Stop"  # so depois do PyInstaller: no PowerShell 5.1 o INFO dele no stderr vira erro

$versao = & "$raiz\.venv\Scripts\python.exe" -c "import sys; sys.path.insert(0, 'src'); from versao import VERSAO; print(VERSAO)"
$pacote = "$raiz\build\pacote"
if (Test-Path $pacote) { Remove-Item $pacote -Recurse -Force }
New-Item -ItemType Directory -Force "$pacote\planilhas", "$raiz\release" | Out-Null
Copy-Item "$raiz\build\dist\RPA-SRP.exe" $pacote
Copy-Item "$raiz\vendor\tesseract" "$pacote\tesseract" -Recurse
Copy-Item "$raiz\planilhas\BASE PADRAO.xlsx" "$pacote\planilhas"
Copy-Item "$raiz\scripts\instalar.cmd", "$raiz\scripts\instalar.ps1" $pacote
$zip = "$raiz\release\RPA-SRP-v$versao"
# so a versao atual fica na pasta: as anteriores estao no GitHub Releases
Get-ChildItem "$raiz\release\RPA-SRP-v*.zip" | Where-Object { $_.Name -ne "RPA-SRP-v$versao.zip" } | Remove-Item -Force
# shutil em vez de Compress-Archive: o do PowerShell 5.1 grava caminhos com "\" (outros descompactadores reclamam)
& "$raiz\.venv\Scripts\python.exe" -c "import shutil, sys; shutil.make_archive(sys.argv[1], 'zip', sys.argv[2])" $zip $pacote
Write-Host "Pacote: $zip.zip"
