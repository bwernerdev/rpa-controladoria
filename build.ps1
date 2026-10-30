# Gera dist\RPA-SRP.exe (janela, arquivo unico). Rode na pasta RPA.
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --onefile --windowed --name RPA-SRP --paths src --collect-all tkinterdnd2 --distpath dist --workpath build --specpath build src\gui.py
