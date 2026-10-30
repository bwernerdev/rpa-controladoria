# vendor
`tesseract/` (versionado, ~24 MB) guarda o Tesseract OCR portatil usado para conferir a Quantidade e ler o cabecalho e os avisos na tela.
Origem: `https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe`
(SHA-256 `c885fff6998e0608ba4bb8ab51436e1c6775c2bafc2559a19b423e18678b60c9`, 50175248 bytes). Nunca executar o instalador: extrair com `7z x` e copiar `tesseract.exe`, as `*.dll`, `tessdata/eng.traineddata` e `tessdata/configs/`.
Depois de copiar, enxugar (163,6 MB -> 24,5 MB; o limite por arquivo do GitHub e 100 MB) e rodar os testes:
1. Tirar a depuracao dos binarios (`libtesseract-5.dll` 96,8 -> 3,1 MB): `.\.venv\Scripts\python.exe scripts\enxugar_dll.py (Get-ChildItem vendor\tesseract\*.dll, vendor\tesseract\*.exe)`.
2. Apagar as DLLs que o `tesseract.exe` nao carrega (ICU, cairo, pango, glib...: so servem ao treino): `.\.venv\Scripts\python.exe scripts\dlls_usadas.py vendor\tesseract` (as marcadas "sobra").
Rodando pelo codigo-fonte (`src\gui.py`), o app usa esta pasta direto. Para o exe, o `build.ps1` monta `release\RPA-SRP-vX.Y.Z.zip` com o exe e esta pasta como `tesseract\` ao lado. Sem ela o app roda, mas a conferencia e a deteccao do C/Custo pelo cabecalho ficam desativadas com aviso.
