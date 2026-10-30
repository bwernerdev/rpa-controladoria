# Gera dist\RPA-SRP.exe (janela, arquivo unico). Rode na pasta RPA.
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --distpath dist --workpath build --specpath build src\gui.py
# Tesseract portatil ao lado do exe (conferencia da Quantidade por OCR)
if (Test-Path vendor\tesseract) { Copy-Item vendor\tesseract dist\tesseract -Recurse -Force }
