# Publica a versao atual (src\versao.py) no GitHub Releases. Rode na pasta RPA, depois de .\build.ps1 e do commit.
# Precisa: ferramentas\gh (gh portatil, fora do git) logado (ver CLAUDE.md), scripts\notas\vX.Y.Z.md e release\RPA-SRP-vX.Y.Z.zip.
# Sem $ErrorActionPreference = "Stop": no PowerShell 5.1 o git/gh escrevem progresso no stderr e isso viraria erro.
# Cada passo confere $LASTEXITCODE.
$raiz = $PWD.Path
$env:GH_CONFIG_DIR = "$raiz\ferramentas\gh-config"
$gh = "$raiz\ferramentas\gh\bin\gh.exe"
$repo = "bwernerdev/rpa-controladoria"
$versao = & "$raiz\.venv\Scripts\python.exe" -c "import sys; sys.path.insert(0, 'src'); from versao import VERSAO; print(VERSAO)"
$tag = "v$versao"
$zip = "$raiz\release\RPA-SRP-$tag.zip"
$notas = "$raiz\scripts\notas\$tag.md"

if (-not (Test-Path $gh)) { throw "gh nao encontrado em $gh" }
if (-not (Test-Path $zip)) { throw "Falta $zip : rode .\build.ps1" }
if (-not (Test-Path $notas)) { throw "Falta $notas (o que mudou nesta versao)" }
if (git status --porcelain) { throw "Ha mudancas sem commit; faca o commit antes de publicar." }
if ((Get-Item $zip).LastWriteTime -lt (Get-Item "$raiz\src\versao.py").LastWriteTime) { throw "O zip e mais antigo que src\versao.py: rode .\build.ps1" }

git push origin HEAD
if ($LASTEXITCODE -ne 0) { throw "git push falhou" }
& $gh release create $tag $zip --repo $repo --target (git rev-parse HEAD) --title $tag --notes-file $notas
if ($LASTEXITCODE -ne 0) { throw "gh release create falhou (a tag $tag ja existe?)" }
Write-Host "Publicado: https://github.com/$repo/releases/tag/$tag"
