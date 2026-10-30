# Gera RPA-SRP.exe (arquivo unico, com tela de abertura, icone e versao) na raiz do projeto. Rode na pasta RPA.
# Feche o RPA-SRP antes. O exe acha o Tesseract em vendor\tesseract (ou numa pasta tesseract ao lado dele).
# Abertura medida (30/10/2026): ~4,6 s, quase tudo antes do Python comecar (verificacao do Windows/antivirus);
# o formato de pasta (--onedir) deu o mesmo tempo, por isso ficou o arquivo unico.
$raiz = $PWD.Path
& "$raiz\.venv\Scripts\python.exe" src\icone.py
& "$raiz\.venv\Scripts\python.exe" scripts\gerar_versao.py
& "$raiz\.venv\Scripts\python.exe" -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --splash "$raiz\assets\splash.png" --icon "$raiz\assets\icone.ico" --version-file "$raiz\build\version.txt" --distpath "$raiz" --workpath build --specpath build src\gui.py
