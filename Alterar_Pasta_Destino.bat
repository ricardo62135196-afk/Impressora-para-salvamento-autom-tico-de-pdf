@echo off
title Alterar pasta de destino - Impressora PDF Virtual
color 0B

set "APP_FOLDER=%ProgramData%\PDFVirtualPrinter"
set "CONFIG_FILE=%APP_FOLDER%\config.ini"

if not exist "%APP_FOLDER%" (
    echo [ERRO] A Impressora PDF Virtual ainda nao foi instalada.
    echo        Execute primeiro o Instalar_Impressora_PDF.bat
    pause
    exit /b 1
)

cls
echo ================================================================
echo    ALTERAR PASTA DE DESTINO DOS PDFs
echo ================================================================
echo.
echo Escolha a nova pasta onde os PDFs serao salvos...
echo.

for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Windows.Forms;" ^
    "$f = New-Object System.Windows.Forms.FolderBrowserDialog;" ^
    "$f.Description = 'Selecione a nova pasta onde os PDFs serao salvos';" ^
    "$f.ShowNewFolderButton = $true;" ^
    "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.SelectedPath } else { Write-Output '' }"`) do set "DEST_FOLDER=%%F"

if "%DEST_FOLDER%"=="" (
    echo [AVISO] Nenhuma pasta selecionada. Nada foi alterado.
    pause
    exit /b 0
)

(
echo SaveFolder=%DEST_FOLDER%
) > "%CONFIG_FILE%"

echo.
echo [OK] Nova pasta de destino: %DEST_FOLDER%
echo      O monitor ja usara essa pasta automaticamente no proximo trabalho de impressao.
echo.
pause
