# vendor
`tesseract/` (ignorado pelo git, ~164 MB) guarda o Tesseract OCR portatil usado para conferir a Quantidade na tela.
Origem: `https://github.com/UB-Mannheim/tesseract/releases/download/v5.4.0.20240606/tesseract-ocr-w64-setup-5.4.0.20240606.exe`
(SHA-256 `c885fff6998e0608ba4bb8ab51436e1c6775c2bafc2559a19b423e18678b60c9`, 50175248 bytes). Nunca executar o instalador: extrair com `7z x` e copiar `tesseract.exe`, todas as `*.dll`, `tessdata/eng.traineddata` e `tessdata/configs/`.
Para distribuir, a pasta `tesseract/` vai ao lado do `RPA-SRP.exe` (o build.ps1 copia para `dist\tesseract`). Sem ela o app roda, mas a conferencia fica desativada com aviso.
