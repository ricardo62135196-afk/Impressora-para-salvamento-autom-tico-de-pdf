<#
    Instalador / Painel - Impressora PDF Virtual
    Interface unica reunindo: Instalar impressora + Alterar pasta de destino
#>

param(
    [ValidateSet('Install','AlterarPasta')]
    [string]$Action
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$APP_FOLDER     = "$env:ProgramData\PDFVirtualPrinter"
$SPOOL_FOLDER   = "$APP_FOLDER\spool"
$PORT_FILE      = "$SPOOL_FOLDER\job.ps"
$CONFIG_FILE    = "$APP_FOLDER\config.ini"
$MONITOR_SCRIPT = "$APP_FOLDER\monitor_pdf.ps1"
$PRINTER_NAME   = "Impressora PDF Virtual"

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Find-Ghostscript {
    $gs = Get-Command gswin64c -ErrorAction SilentlyContinue
    if ($gs) { return $gs.Source }
    $candidatos = @(
        "C:\Program Files\gs\gs*\bin\gswin64c.exe",
        "C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe"
    )
    foreach ($padrao in $candidatos) {
        $achado = Get-ChildItem -Path $padrao -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($achado) { return $achado.FullName }
    }
    return $null
}

function Install-Printer {
    Write-Host "================================================================"
    Write-Host "   INSTALADOR - IMPRESSORA VIRTUAL PDF"
    Write-Host "================================================================"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Host "[ERRO] Este processo precisa ser executado como Administrador." -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        return
    }

    Write-Host "[1/9] Criando pastas de trabalho..."
    New-Item -ItemType Directory -Path $APP_FOLDER -Force | Out-Null
    New-Item -ItemType Directory -Path $SPOOL_FOLDER -Force | Out-Null
    Write-Host "      Liberando permissao de escrita para qualquer usuario..."
    icacls $APP_FOLDER /grant *S-1-5-32-545:(OI)(CI)M /T | Out-Null
    Write-Host "      [OK] Pastas criadas em: $APP_FOLDER" -ForegroundColor Green
    Write-Host ""

    Write-Host "[2/9] Verificando Ghostscript..."
    $GS_PATH = Find-Ghostscript
    if (-not $GS_PATH) {
        Write-Host "      [ERRO] Ghostscript NAO encontrado!" -ForegroundColor Red
        Write-Host "      Baixe em: https://ghostscript.com/releases/gsdnld.html"
        Read-Host "Pressione Enter para sair"
        return
    }
    Write-Host "      [OK] Ghostscript encontrado: $GS_PATH" -ForegroundColor Green
    Write-Host ""

    Write-Host "[3/9] Escolha a pasta onde os PDFs serao salvos..."
    $picker = New-Object System.Windows.Forms.FolderBrowserDialog
    $picker.Description = "Selecione a pasta onde os PDFs impressos serao salvos"
    $picker.ShowNewFolderButton = $true
    $DEST_FOLDER = $null
    if ($picker.ShowDialog() -eq 'OK') { $DEST_FOLDER = $picker.SelectedPath }
    if (-not $DEST_FOLDER) {
        $DEST_FOLDER = "$env:USERPROFILE\Documents\PDFs"
        New-Item -ItemType Directory -Path $DEST_FOLDER -Force | Out-Null
        Write-Host "      [AVISO] Nenhuma pasta selecionada. Usando pasta padrao." -ForegroundColor Yellow
    }
    Write-Host "      [OK] PDFs serao salvos em: $DEST_FOLDER" -ForegroundColor Green
    Set-Content -Path $CONFIG_FILE -Value "SaveFolder=$DEST_FOLDER" -Encoding ASCII
    Write-Host ""

    Write-Host "[4/9] Limpando instalacao anterior, se houver..."
    Remove-Printer -Name $PRINTER_NAME -ErrorAction SilentlyContinue
    Remove-PrinterPort -Name $PORT_FILE -ErrorAction SilentlyContinue
    Remove-Item $PORT_FILE -Force -ErrorAction SilentlyContinue
    Write-Host "      [OK] Limpeza concluida" -ForegroundColor Green
    Write-Host ""

    Write-Host "[5/9] Criando porta local..."
    Add-PrinterPort -Name $PORT_FILE -ErrorAction SilentlyContinue
    Write-Host "      [OK] Porta criada: $PORT_FILE" -ForegroundColor Green
    Write-Host ""

    Write-Host "[6/9] Instalando impressora..."
    rundll32 printui.dll,PrintUIEntry /if /b "$PRINTER_NAME" /f "$env:windir\inf\ntprint.inf" /r "$PORT_FILE" /m "MS Publisher Imagesetter"
    Start-Sleep -Seconds 1
    if (-not (Get-Printer -Name $PRINTER_NAME -ErrorAction SilentlyContinue)) {
        Write-Host "      [AVISO] Tentando driver alternativo..." -ForegroundColor Yellow
        rundll32 printui.dll,PrintUIEntry /if /b "$PRINTER_NAME" /f "$env:windir\inf\ntprint.inf" /r "$PORT_FILE" /m "MS Publisher Color Printer"
    }
    Write-Host "      [OK] Impressora instalada" -ForegroundColor Green
    Write-Host ""

    Write-Host "[7/9] Definindo como impressora padrao..."
    $p = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$PRINTER_NAME'" -ErrorAction SilentlyContinue
    if ($p) { Invoke-CimMethod -InputObject $p -MethodName SetDefaultPrinter | Out-Null }
    Write-Host "      [OK] Impressora definida como padrao" -ForegroundColor Green
    Write-Host ""

    Write-Host "[8/9] Criando monitor de conversao..."
    $monitorContent = @"
# Monitor da Impressora Virtual PDF
`$ErrorActionPreference = "SilentlyContinue"

`$PortFile   = "$PORT_FILE"
`$ConfigFile = "$CONFIG_FILE"
`$GS         = "$GS_PATH"

Write-Host "=== MONITOR IMPRESSORA PDF VIRTUAL ===" -ForegroundColor Green
Write-Host "Arquivo de porta: `$PortFile" -ForegroundColor Yellow
Write-Host "Config: `$ConfigFile" -ForegroundColor Yellow

function Get-DestFolder {
    `$linha = Get-Content `$ConfigFile | Where-Object { `$_ -like "SaveFolder=*" }
    if (-not `$linha) { return "`$env:USERPROFILE\Documents\PDFs" }
    `$folder = `$linha -replace "^SaveFolder=", ""
    if (-not (Test-Path `$folder)) { New-Item -ItemType Directory -Path `$folder -Force | Out-Null }
    return `$folder
}

function Wait-FileReady {
    param(`$Path)
    `$pronto = `$false
    while (-not `$pronto) {
        Start-Sleep -Milliseconds 800
        try {
            `$s = [System.IO.File]::Open(`$Path, 'Open', 'ReadWrite', 'None')
            `$s.Close()
            `$pronto = `$true
        } catch {
            `$pronto = `$false
        }
    }
}

Write-Host "Monitor ativo. Aguardando impressoes..." -ForegroundColor Green

while (`$true) {
    if (Test-Path `$PortFile) {
        Wait-FileReady -Path `$PortFile
        Start-Sleep -Milliseconds 500

        `$destFolder = Get-DestFolder
        `$timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        `$pdfPath    = Join-Path `$destFolder "Documento_`$timestamp.pdf"

        `$gsArgs = @(
            "-dNOPAUSE", "-dBATCH", "-dSAFER",
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.7",
            "-dPDFSETTINGS=/prepress",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dAutoRotatePages=/None",
            "-sOutputFile=`$pdfPath",
            "`$PortFile"
        )

        try {
            Start-Process -FilePath `$GS -ArgumentList `$gsArgs -NoNewWindow -Wait
        } catch {
            Write-Host "[ERRO] Falha ao converter: `$_" -ForegroundColor Red
        }

        if ((Test-Path `$pdfPath) -and ((Get-Item `$pdfPath).Length -gt 1024)) {
            Write-Host "[OK] PDF salvo: `$pdfPath" -ForegroundColor Green
        } else {
            Write-Host "[ERRO] PDF nao foi gerado corretamente" -ForegroundColor Red
            Remove-Item `$pdfPath -Force -ErrorAction SilentlyContinue
        }

        Remove-Item `$PortFile -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 1
}
"@
    Set-Content -Path $MONITOR_SCRIPT -Value $monitorContent -Encoding UTF8
    Write-Host "      [OK] Monitor criado" -ForegroundColor Green
    Write-Host ""

    Write-Host "[9/9] Configurando inicializacao automatica do monitor..."
    $wsh = New-Object -ComObject WScript.Shell
    $startupFolder = $wsh.SpecialFolders('AllUsersStartup')
    if (-not (Test-Path $startupFolder)) { New-Item -ItemType Directory -Path $startupFolder -Force | Out-Null }
    $shortcutPath = Join-Path $startupFolder 'Monitor PDF Virtual.lnk'
    $s = $wsh.CreateShortcut($shortcutPath)
    $s.TargetPath = "powershell.exe"
    $s.Arguments = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$MONITOR_SCRIPT`""
    $s.WorkingDirectory = $APP_FOLDER
    $s.Save()
    Write-Host "      [OK] Atalho criado para todos os usuarios" -ForegroundColor Green
    Write-Host ""

    Start-Process powershell.exe -ArgumentList "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$MONITOR_SCRIPT`"" -WorkingDirectory $APP_FOLDER

    Write-Host "================================================================"
    Write-Host "   INSTALACAO CONCLUIDA!"
    Write-Host "================================================================"
    Write-Host ""
    Write-Host "Impressora padrao: $PRINTER_NAME"
    Write-Host "Pasta onde os PDFs sao salvos: $DEST_FOLDER"
    Write-Host "Ghostscript usado: $GS_PATH"
    Write-Host ""
    Read-Host "Pressione Enter para fechar"
}

function Set-DestinationFolder {
    Write-Host "================================================================"
    Write-Host "   ALTERAR PASTA DE DESTINO DOS PDFs"
    Write-Host "================================================================"
    Write-Host ""

    if (-not (Test-Path $APP_FOLDER)) {
        Write-Host "[ERRO] A Impressora PDF Virtual ainda nao foi instalada." -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        return
    }

    $pastaAtual = $null
    if (Test-Path $CONFIG_FILE) {
        $linha = Get-Content $CONFIG_FILE | Where-Object { $_ -like "SaveFolder=*" }
        if ($linha) { $pastaAtual = $linha -replace "^SaveFolder=", "" }
    }
    if ($pastaAtual) {
        Write-Host "Pasta atual configurada: $pastaAtual"
    } else {
        Write-Host "Nenhuma pasta configurada ainda."
    }
    Write-Host ""

    $picker = New-Object System.Windows.Forms.FolderBrowserDialog
    $picker.Description = "Selecione a nova pasta onde os PDFs serao salvos"
    $picker.ShowNewFolderButton = $true
    if ($pastaAtual -and (Test-Path $pastaAtual)) { $picker.SelectedPath = $pastaAtual }

    if ($picker.ShowDialog() -ne 'OK') {
        Write-Host "[AVISO] Nenhuma pasta selecionada. Nada foi alterado." -ForegroundColor Yellow
        Read-Host "Pressione Enter para sair"
        return
    }

    $novaPasta = $picker.SelectedPath
    try {
        Set-Content -Path $CONFIG_FILE -Value "SaveFolder=$novaPasta" -Encoding ASCII -ErrorAction Stop
        Write-Host "[OK] config.ini atualizado com sucesso" -ForegroundColor Green
    } catch {
        Write-Host "[ERRO] Nao foi possivel gravar o config.ini: $_" -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        return
    }

    Write-Host ""
    Write-Host "Conferindo o que foi gravado:"
    Get-Content $CONFIG_FILE
    Write-Host ""
    Write-Host "[OK] Nova pasta de destino: $novaPasta" -ForegroundColor Green
    Write-Host "     O monitor ja usara essa pasta automaticamente no proximo trabalho de impressao."
    Write-Host ""
    Read-Host "Pressione Enter para fechar"
}

# --- Execucao direta de uma acao especifica (usado apos auto-elevacao) ---
if ($Action -eq 'Install') { Install-Printer; exit }
if ($Action -eq 'AlterarPasta') { Set-DestinationFolder; exit }

# --- Interface grafica principal ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Impressora PDF Virtual"
$form.Size = New-Object System.Drawing.Size(420, 260)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$titulo = New-Object System.Windows.Forms.Label
$titulo.Text = "Impressora PDF Virtual"
$titulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titulo.AutoSize = $true
$titulo.Location = New-Object System.Drawing.Point(20, 20)
$form.Controls.Add($titulo)

$subtitulo = New-Object System.Windows.Forms.Label
$subtitulo.Text = "Escolha uma opcao:"
$subtitulo.AutoSize = $true
$subtitulo.Location = New-Object System.Drawing.Point(20, 55)
$form.Controls.Add($subtitulo)

$btnInstalar = New-Object System.Windows.Forms.Button
$btnInstalar.Text = "Instalar / Reinstalar Impressora"
$btnInstalar.Size = New-Object System.Drawing.Size(360, 45)
$btnInstalar.Location = New-Object System.Drawing.Point(20, 90)
$btnInstalar.Add_Click({
    $form.Hide()
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Install" -Wait
    $form.Show()
})
$form.Controls.Add($btnInstalar)

$btnAlterar = New-Object System.Windows.Forms.Button
$btnAlterar.Text = "Alterar Pasta de Destino"
$btnAlterar.Size = New-Object System.Drawing.Size(360, 45)
$btnAlterar.Location = New-Object System.Drawing.Point(20, 145)
$btnAlterar.Add_Click({
    $form.Hide()
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action AlterarPasta" -Wait
    $form.Show()
})
$form.Controls.Add($btnAlterar)

$btnSair = New-Object System.Windows.Forms.Button
$btnSair.Text = "Sair"
$btnSair.Size = New-Object System.Drawing.Size(360, 30)
$btnSair.Location = New-Object System.Drawing.Point(20, 195)
$btnSair.Add_Click({ $form.Close() })
$form.Controls.Add($btnSair)

[void]$form.ShowDialog()
