# Gera dist\RPA-SRP.exe (janela, arquivo unico, com icone e versao). Rode na pasta RPA. Feche o RPA-SRP antes.
$raiz = $PWD.Path
& "$raiz\.venv\Scripts\python.exe" src\icone.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_versao.py
& "$raiz\.venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --icon "$raiz\assets\icone.ico" --version-file "$raiz\build\version.txt" --distpath dist --workpath build --specpath build src\gui.py
# Tesseract portatil ao lado do exe (conferencia da Quantidade por OCR)
if (Test-Path vendor\tesseract) { Copy-Item vendor\tesseract dist\tesseract -Recurse -Force }
