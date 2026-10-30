@echo off
rem Duplo clique: instala ou atualiza o Robo SRP (sem administrador) e cria o atalho na area de trabalho.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar.ps1"
if errorlevel 1 (
    echo.
    echo A instalacao falhou. Leia a mensagem acima.
)
echo.
pause
