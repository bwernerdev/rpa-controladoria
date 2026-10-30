# Gera RPA-SRP.exe na raiz do projeto (janela, arquivo unico, com icone e versao). Rode na pasta RPA. Feche o RPA-SRP antes.
# O exe acha o Tesseract em vendor\tesseract (ou numa pasta tesseract ao lado dele, ao levar para outra maquina).
$raiz = $PWD.Path
& "$raiz\.venv\Scripts\python.exe" src\icone.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_versao.py
& "$raiz\.venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --icon "$raiz\assets\icone.ico" --version-file "$raiz\build\version.txt" --distpath "$raiz" --workpath build --specpath build src\gui.py
