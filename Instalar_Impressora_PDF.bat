@echo off
title Instalador - Impressora Virtual PDF
color 0A
mode con: cols=100 lines=40

:: ============================================================
:: Verificar se esta rodando como administrador
:: ============================================================
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo [ERRO] Execute este arquivo como ADMINISTRADOR!
    echo        Clique com o botao direito e escolha "Executar como administrador".
    echo.
    pause
    exit /b 1
)

cls
echo ================================================================
echo    INSTALADOR - IMPRESSORA VIRTUAL PDF
echo    Cria impressora padrao que salva PDFs em pasta escolhida
echo ================================================================
echo.

:: ============================================================
:: Configuracoes gerais
:: ============================================================
set "PRINTER_NAME=Impressora PDF Virtual"
set "PORT_NAME=PDF_VIRTUAL_PORT:"
set "APP_FOLDER=%ProgramData%\PDFVirtualPrinter"
set "SPOOL_FOLDER=%APP_FOLDER%\spool"
set "PORT_FILE=%SPOOL_FOLDER%\job.ps"
set "CONFIG_FILE=%APP_FOLDER%\config.ini"
set "MONITOR_SCRIPT=%APP_FOLDER%\monitor_pdf.ps1"

echo [1/9] Criando pastas de trabalho...
if not exist "%APP_FOLDER%" mkdir "%APP_FOLDER%"
if not exist "%SPOOL_FOLDER%" mkdir "%SPOOL_FOLDER%"
echo       Liberando permissao de escrita para qualquer usuario...
icacls "%APP_FOLDER%" /grant *S-1-5-32-545:(OI)(CI)M /T >nul 2>&1
echo       [OK] Pastas criadas em: %APP_FOLDER%
echo.

:: ============================================================
:: Verificar Ghostscript
:: ============================================================
echo [2/9] Verificando Ghostscript...
set "GS_PATH="
where gswin64c >nul 2>&1
if %errorLevel% equ 0 (
    set "GS_PATH=gswin64c"
) else (
    for /d %%g in ("C:\Program Files\gs\gs*") do (
        if exist "%%g\bin\gswin64c.exe" set "GS_PATH=%%g\bin\gswin64c.exe"
    )
    if not defined GS_PATH (
        for /d %%g in ("C:\Program Files (x86)\gs\gs*") do (
            if exist "%%g\bin\gswin32c.exe" set "GS_PATH=%%g\bin\gswin32c.exe"
        )
    )
)

if not defined GS_PATH (
    echo       [ERRO] Ghostscript NAO encontrado!
    echo       Baixe em: https://ghostscript.com/releases/gsdnld.html
    pause
    exit /b 1
)
echo       [OK] Ghostscript encontrado: %GS_PATH%
echo.

:: ============================================================
:: Escolher a pasta onde os PDFs serao salvos
:: ============================================================
echo [3/9] Escolha a pasta onde os PDFs serao salvos...
echo       (uma janela sera aberta - se nao aparecer, verifique a barra de tarefas)
echo.

for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
    "Add-Type -AssemblyName System.Windows.Forms;" ^
    "$f = New-Object System.Windows.Forms.FolderBrowserDialog;" ^
    "$f.Description = 'Selecione a pasta onde os PDFs impressos serao salvos';" ^
    "$f.ShowNewFolderButton = $true;" ^
    "if ($f.ShowDialog() -eq 'OK') { Write-Output $f.SelectedPath } else { Write-Output '' }"`) do set "DEST_FOLDER=%%F"

if "%DEST_FOLDER%"=="" (
    echo       [AVISO] Nenhuma pasta selecionada. Usando pasta padrao.
    set "DEST_FOLDER=%USERPROFILE%\Documents\PDFs"
    if not exist "%DEST_FOLDER%" mkdir "%DEST_FOLDER%"
)

echo       [OK] PDFs serao salvos em: %DEST_FOLDER%
echo.

:: Salvar escolha no arquivo de configuracao
(
echo SaveFolder=%DEST_FOLDER%
) > "%CONFIG_FILE%"

:: ============================================================
:: Remover impressora/porta antigas (reinstalacao limpa)
:: ============================================================
echo [4/9] Limpando instalacao anterior, se houver...
powershell -NoProfile -Command "Remove-Printer -Name '%PRINTER_NAME%' -ErrorAction SilentlyContinue"
powershell -NoProfile -Command "Remove-PrinterPort -Name '%PORT_FILE%' -ErrorAction SilentlyContinue"
if exist "%PORT_FILE%" del /f /q "%PORT_FILE%" >nul 2>&1
echo       [OK] Limpeza concluida
echo.

:: ============================================================
:: Criar porta local apontando para um arquivo fixo
:: (a impressora sempre escreve o trabalho neste mesmo arquivo)
:: ============================================================
echo [5/9] Criando porta local...
powershell -NoProfile -Command "Add-PrinterPort -Name '%PORT_FILE%' -ErrorAction SilentlyContinue"
echo       [OK] Porta criada: %PORT_FILE%
echo.

:: ============================================================
:: Instalar a impressora com driver PostScript
:: ============================================================
echo [6/9] Instalando impressora...
rundll32 printui.dll,PrintUIEntry /if /b "%PRINTER_NAME%" /f "%windir%\inf\ntprint.inf" /r "%PORT_FILE%" /m "MS Publisher Imagesetter"

powershell -NoProfile -Command "if (-not (Get-Printer -Name '%PRINTER_NAME%' -ErrorAction SilentlyContinue)) { exit 1 } else { exit 0 }"
if %errorLevel% neq 0 (
    echo       [AVISO] Tentando driver alternativo...
    rundll32 printui.dll,PrintUIEntry /if /b "%PRINTER_NAME%" /f "%windir%\inf\ntprint.inf" /r "%PORT_FILE%" /m "MS Publisher Color Printer"
)
echo       [OK] Impressora instalada
echo.

:: ============================================================
:: Definir como impressora padrao
:: ============================================================
echo [7/9] Definindo como impressora padrao...
powershell -NoProfile -Command ^
    "$p = Get-CimInstance -ClassName Win32_Printer -Filter \"Name='%PRINTER_NAME%'\";" ^
    "if ($p) { Invoke-CimMethod -InputObject $p -MethodName SetDefaultPrinter | Out-Null }"
echo       [OK] "%PRINTER_NAME%" definida como padrao
echo.

:: ============================================================
:: Criar o script de monitoramento (converte para PDF)
:: ============================================================
echo [8/9] Criando monitor de conversao...
(
echo # Monitor da Impressora Virtual PDF
echo $ErrorActionPreference = "SilentlyContinue"
echo.
echo $PortFile   = "%PORT_FILE%"
echo $ConfigFile = "%CONFIG_FILE%"
echo $GS         = "%GS_PATH%"
echo.
echo Write-Host "=== MONITOR IMPRESSORA PDF VIRTUAL ===" -ForegroundColor Green
echo Write-Host "Arquivo de porta: $PortFile" -ForegroundColor Yellow
echo Write-Host "Config: $ConfigFile" -ForegroundColor Yellow
echo.
echo function Get-DestFolder {
echo     $linha = Get-Content $ConfigFile ^| Where-Object { $_ -like "SaveFolder=*" }
echo     if ^(-not $linha^) { return "$env:USERPROFILE\Documents\PDFs" }
echo     $folder = $linha -replace "^SaveFolder=", ""
echo     if ^(-not ^(Test-Path $folder^)^) { New-Item -ItemType Directory -Path $folder -Force ^| Out-Null }
echo     return $folder
echo }
echo.
echo function Wait-FileReady {
echo     param^($Path^)
echo     $pronto = $false
echo     while ^(-not $pronto^) {
echo         Start-Sleep -Milliseconds 800
echo         try {
echo             $s = [System.IO.File]::Open^($Path, 'Open', 'ReadWrite', 'None'^)
echo             $s.Close^(^)
echo             $pronto = $true
echo         } catch {
echo             $pronto = $false
echo         }
echo     }
echo }
echo.
echo Write-Host "Monitor ativo. Aguardando impressoes..." -ForegroundColor Green
echo.
echo while ^($true^) {
echo     if ^(Test-Path $PortFile^) {
echo         Wait-FileReady -Path $PortFile
echo         Start-Sleep -Milliseconds 500
echo.
echo         $destFolder = Get-DestFolder
echo         $timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
echo         $pdfPath    = Join-Path $destFolder "Documento_$timestamp.pdf"
echo.
echo         $gsArgs = @^(
echo             "-dNOPAUSE", "-dBATCH", "-dSAFER",
echo             "-sDEVICE=pdfwrite",
echo             "-dCompatibilityLevel=1.7",
echo             "-dPDFSETTINGS=/prepress",
echo             "-dEmbedAllFonts=true",
echo             "-dSubsetFonts=true",
echo             "-dAutoRotatePages=/None",
echo             "-sOutputFile=$pdfPath",
echo             "$PortFile"
echo         ^)
echo.
echo         try {
echo             Start-Process -FilePath $GS -ArgumentList $gsArgs -NoNewWindow -Wait
echo         } catch {
echo             Write-Host "[ERRO] Falha ao converter: $_" -ForegroundColor Red
echo         }
echo.
echo         if ^(^(Test-Path $pdfPath^) -and ^(^(Get-Item $pdfPath^).Length -gt 1024^)^) {
echo             Write-Host "[OK] PDF salvo: $pdfPath" -ForegroundColor Green
echo         } else {
echo             Write-Host "[ERRO] PDF nao foi gerado corretamente" -ForegroundColor Red
echo             Remove-Item $pdfPath -Force -ErrorAction SilentlyContinue
echo         }
echo.
echo         Remove-Item $PortFile -Force -ErrorAction SilentlyContinue
echo     }
echo     Start-Sleep -Seconds 1
echo }
) > "%MONITOR_SCRIPT%"
echo       [OK] Monitor criado
echo.

:: ============================================================
:: Fazer o monitor iniciar automaticamente com o Windows
:: ============================================================
echo [9/9] Configurando inicializacao automatica do monitor...
powershell -NoProfile -Command ^
    "$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument '-WindowStyle Hidden -ExecutionPolicy Bypass -File \"%MONITOR_SCRIPT%\"' -WorkingDirectory '%APP_FOLDER%';" ^
    "$trigger = New-ScheduledTaskTrigger -AtLogOn;" ^
    "$principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -RunLevel Limited;" ^
    "Unregister-ScheduledTask -TaskName 'Monitor PDF Virtual' -Confirm:$false -ErrorAction SilentlyContinue;" ^
    "Register-ScheduledTask -TaskName 'Monitor PDF Virtual' -Action $action -Trigger $trigger -Principal $principal -Description 'Monitor de conversao de PDF' -Force | Out-Null;" ^
    "Write-Host 'Tarefa agendada criada: Monitor PDF Virtual (roda para qualquer usuario)'"
echo       [OK] Monitor sera iniciado automaticamente ao ligar o PC
echo.

:: Iniciar o monitor agora tambem
start /min powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%MONITOR_SCRIPT%"

echo ================================================================
echo    INSTALACAO CONCLUIDA!
echo ================================================================
echo.
echo Impressora padrao: %PRINTER_NAME%
echo Pasta onde os PDFs sao salvos: %DEST_FOLDER%
echo Ghostscript usado: %GS_PATH%
echo.
echo Para IMPRIMIR qualquer arquivo, basta usar a opcao "Imprimir" normal
echo do programa (Word, navegador, etc.) - a impressora ja esta como padrao.
echo O PDF aparecera automaticamente na pasta escolhida em alguns segundos.
echo.
echo Para TROCAR a pasta de destino no futuro, use o arquivo:
echo    Alterar_Pasta_Destino.bat
echo.
pause
