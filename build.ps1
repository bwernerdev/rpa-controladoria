# Gera RPA-SRP.exe (arquivo unico, com tela de abertura, icone e versao) na raiz do projeto. Rode na pasta RPA.
# Feche o RPA-SRP antes. O exe acha o Tesseract em vendor\tesseract (ou numa pasta tesseract ao lado dele).
# Abertura medida (30/10/2026): ~4,6 s, quase tudo antes do Python comecar (verificacao do Windows/antivirus);
# o formato de pasta (--onedir) deu o mesmo tempo, por isso ficou o arquivo unico.
# Depois do exe, monta release\RPA-SRP-vX.Y.Z.zip (exe + tesseract\ + BASE PADRAO.xlsx + instalar.cmd) para o GitHub Releases;
# o exe e o zip nao vao ao git.
$raiz = $PWD.Path
& "$raiz\.venv\Scripts\python.exe" src\icone.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_versao.py
& "$raiz\.venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --splash "$raiz\assets\splash.png" --icon "$raiz\assets\icone.ico" --version-file "$raiz\build\version.txt" --distpath "$raiz" --workpath build --specpath build src\gui.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller falhou" }
$ErrorActionPreference = "Stop"  # so depois do PyInstaller: no PowerShell 5.1 o INFO dele no stderr vira erro

$versao = & "$raiz\.venv\Scripts\python.exe" -c "import sys; sys.path.insert(0, 'src'); from versao import VERSAO; print(VERSAO)"
$pacote = "$raiz\build\pacote"
if (Test-Path $pacote) { Remove-Item $pacote -Recurse -Force }
New-Item -ItemType Directory -Force "$pacote\planilhas", "$raiz\release" | Out-Null
Copy-Item "$raiz\RPA-SRP.exe" $pacote
Copy-Item "$raiz\vendor\tesseract" "$pacote\tesseract" -Recurse
Copy-Item "$raiz\planilhas\BASE PADRAO.xlsx" "$pacote\planilhas"
Copy-Item "$raiz\scripts\instalar.cmd", "$raiz\scripts\instalar.ps1" $pacote
$zip = "$raiz\release\RPA-SRP-v$versao"
# shutil em vez de Compress-Archive: o do PowerShell 5.1 grava caminhos com "\" (outros descompactadores reclamam)
& "$raiz\.venv\Scripts\python.exe" -c "import shutil, sys; shutil.make_archive(sys.argv[1], 'zip', sys.argv[2])" $zip $pacote
Write-Host "Pacote: $zip.zip"
