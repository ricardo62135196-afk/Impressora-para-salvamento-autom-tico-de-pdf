<#
    Instalador / Painel - Impressora PDF Virtual
    Interface unica reunindo: Instalar impressora + Alterar pasta de destino
#>

param(
    [ValidateSet('Install','AlterarPasta','Uninstall','NamePattern')]
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

function Find-Tesseract {
    $t = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($t) { return $t.Source }
    $candidatos = @(
        "C:\Program Files\Tesseract-OCR\tesseract.exe",
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    )
    foreach ($c in $candidatos) {
        if (Test-Path $c) { return $c }
    }
    return $null
}

function Get-ConfigValue {
    param($Key)
    if (-not (Test-Path $CONFIG_FILE)) { return $null }
    $linha = Get-Content $CONFIG_FILE | Where-Object { $_ -like "$Key=*" }
    if ($linha) { return ($linha -replace "^$Key=", "") }
    return $null
}

function Set-ConfigValue {
    param($Key, $Value)
    $linhas = @()
    if (Test-Path $CONFIG_FILE) {
        $linhas = @(Get-Content $CONFIG_FILE | Where-Object { $_ -notlike "$Key=*" -and $_.Trim() -ne "" })
    }
    $linhas += "$Key=$Value"
    Set-Content -Path $CONFIG_FILE -Value $linhas -Encoding ASCII
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
    icacls $APP_FOLDER /grant "*S-1-5-32-545:(OI)(CI)M" /T | Out-Null
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

    Write-Host "Verificando Tesseract OCR (opcional, usado na nomeacao automatica)..."
    $TESS_PATH = Find-Tesseract
    if ($TESS_PATH) {
        Write-Host "      [OK] Tesseract encontrado: $TESS_PATH" -ForegroundColor Green
    } else {
        Write-Host "      [AVISO] Tesseract nao encontrado. A nomeacao automatica dos PDFs nao vai funcionar" -ForegroundColor Yellow
        Write-Host "               ate instalar em https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Yellow
    }
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
    Set-ConfigValue -Key "SaveFolder" -Value $DEST_FOLDER
    Write-Host ""

    Write-Host "[4/9] Limpando instalacao anterior, se houver..."
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*monitor_pdf.ps1*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Remove-Printer -Name $PRINTER_NAME -ErrorAction SilentlyContinue
    Remove-PrinterPort -Name $PORT_FILE -ErrorAction SilentlyContinue
    Remove-Item $PORT_FILE -Force -ErrorAction SilentlyContinue
    Write-Host "      [OK] Limpeza concluida (monitor antigo parado, impressora e porta removidas)" -ForegroundColor Green
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
`$AppFolder  = "$APP_FOLDER"
`$Tesseract  = "$TESS_PATH"

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

function Get-NamePatterns {
    if (-not (Test-Path `$ConfigFile)) { return @() }
    `$linha = Get-Content `$ConfigFile | Where-Object { `$_ -like "NamePatterns=*" }
    if (-not `$linha) { return @() }
    `$valor = (`$linha -replace "^NamePatterns=", "").Trim()
    if (`$valor -eq "") { return @() }
    return @(`$valor -split '\|\|\|' | ForEach-Object { `$_.Trim() } | Where-Object { `$_ -ne "" })
}

function Remove-Diacritics {
    param(`$Texto)
    `$normalizado = `$Texto.Normalize([Text.NormalizationForm]::FormD)
    `$sb = New-Object System.Text.StringBuilder
    foreach (`$ch in `$normalizado.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory(`$ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]`$sb.Append(`$ch)
        }
    }
    return `$sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Get-FlexPattern {
    param(`$Texto)
    `$chars = `$Texto.ToCharArray() | ForEach-Object { [RegEx]::Escape([string]`$_) }
    return (`$chars -join '\s*')
}

function Get-SanitizedName {
    param(`$Texto)
    `$invalidos = [IO.Path]::GetInvalidFileNameChars() -join ''
    `$regexInvalidos = "[{0}]" -f [RegEx]::Escape(`$invalidos)
    `$limpo = (`$Texto -replace `$regexInvalidos, '_').Trim()
    `$limpo = `$limpo -replace '\s+', ' '
    if (`$limpo.Length -gt 60) { `$limpo = `$limpo.Substring(0, 60).Trim() }
    if (`$limpo -eq "") { return `$null }
    return `$limpo
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
            `$padroes = Get-NamePatterns
            if (`$padroes.Count -gt 0) {
                `$diag = New-Object System.Text.StringBuilder
                [void]`$diag.AppendLine("=== DIAGNOSTICO DE NOMEACAO AUTOMATICA (OCR) ===")
                [void]`$diag.AppendLine("Data/hora: `$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
                [void]`$diag.AppendLine("PDF processado: `$pdfPath")
                [void]`$diag.AppendLine("Rotulos configurados: `$(`$padroes -join ' | ')")
                [void]`$diag.AppendLine("Tesseract: `$Tesseract")
                [void]`$diag.AppendLine("")

                `$partes = @()

                if ([string]::IsNullOrWhiteSpace(`$Tesseract) -or -not (Test-Path `$Tesseract)) {
                    [void]`$diag.AppendLine("[ERRO] Tesseract OCR nao foi encontrado nesta maquina.")
                    [void]`$diag.AppendLine("Instale em https://github.com/UB-Mannheim/tesseract/wiki e reinstale a impressora.")
                } else {
                    `$imgPath = Join-Path `$env:TEMP "gs_ocr_`$timestamp.png"
                    `$imgArgs = @(
                        "-dNOPAUSE", "-dBATCH", "-dSAFER",
                        "-sDEVICE=png16m",
                        "-r300",
                        "-dFirstPage=1",
                        "-dLastPage=1",
                        "-sOutputFile=`$imgPath",
                        "`$pdfPath"
                    )

                    try {
                        `$gsImgResult = Start-Process -FilePath `$GS -ArgumentList `$imgArgs -NoNewWindow -Wait -PassThru
                        [void]`$diag.AppendLine("Renderizacao da pagina 1 em imagem: codigo de saida `$(`$gsImgResult.ExitCode)")
                    } catch {
                        [void]`$diag.AppendLine("[ERRO] Falha ao renderizar pagina em imagem: `$_")
                    }

                    if (Test-Path `$imgPath) {
                        [void]`$diag.AppendLine("Imagem gerada: SIM (`$((Get-Item `$imgPath).Length) bytes)")
                        `$ocrBase = Join-Path `$env:TEMP "gs_ocr_`$timestamp"
                        `$ocrTxtPath = "`${ocrBase}.txt"

                        try {
                            `$tessResult = Start-Process -FilePath `$Tesseract -ArgumentList @("`$imgPath", "`$ocrBase", "-l", "por") -NoNewWindow -Wait -PassThru
                            [void]`$diag.AppendLine("Tesseract rodou. Codigo de saida: `$(`$tessResult.ExitCode)")
                        } catch {
                            [void]`$diag.AppendLine("[ERRO] Falha ao executar o Tesseract: `$_")
                        }

                        if (Test-Path `$ocrTxtPath) {
                            `$textoOriginal = Get-Content `$ocrTxtPath -Raw -ErrorAction SilentlyContinue
                            [void]`$diag.AppendLine("Texto OCR extraido: SIM (`$(if (`$textoOriginal) { `$textoOriginal.Length } else { 0 }) caracteres)")

                            if (`$textoOriginal) {
                                `$texto = Remove-Diacritics -Texto `$textoOriginal
                                `$linhasTexto = `$texto -split "`r`n|`n"

                                foreach (`$rotulo in `$padroes) {
                                    `$rotuloNorm = Remove-Diacritics -Texto `$rotulo
                                    `$flex = Get-FlexPattern -Texto `$rotuloNorm
                                    `$valor = `$null
                                    `$metodo = "nao encontrado"

                                    if (`$texto -match "(?i)`$flex\s*[:\-]?\s*([^\r\n]+)") {
                                        `$candidato = Get-SanitizedName -Texto `$Matches[1]
                                        if (`$candidato) { `$valor = `$candidato; `$metodo = "mesma linha" }
                                    }

                                    if (-not `$valor) {
                                        for (`$i = 0; `$i -lt `$linhasTexto.Count; `$i++) {
                                            if (`$linhasTexto[`$i] -match "(?i)`$flex") {
                                                for (`$j = `$i + 1; `$j -lt `$linhasTexto.Count; `$j++) {
                                                    `$candidataLinha = `$linhasTexto[`$j].Trim()
                                                    if (`$candidataLinha -ne "") {
                                                        `$valor = Get-SanitizedName -Texto `$candidataLinha
                                                        `$metodo = "linha seguinte"
                                                        break
                                                    }
                                                }
                                                break
                                            }
                                        }
                                    }

                                    if (`$valor) {
                                        `$partes += `$valor
                                        [void]`$diag.AppendLine("Rotulo '`$rotulo' -> ENCONTRADO ('`$valor', metodo: `$metodo)")
                                    } else {
                                        [void]`$diag.AppendLine("Rotulo '`$rotulo' -> NAO encontrado no texto OCR")
                                    }
                                }

                                [void]`$diag.AppendLine("")
                                [void]`$diag.AppendLine("--- TEXTO RECONHECIDO PELO OCR (na integra) ---")
                                [void]`$diag.AppendLine(`$textoOriginal)
                            }
                            Remove-Item `$ocrTxtPath -Force -ErrorAction SilentlyContinue
                        } else {
                            [void]`$diag.AppendLine("[AVISO] O Tesseract nao gerou arquivo de texto.")
                        }
                        Remove-Item `$imgPath -Force -ErrorAction SilentlyContinue
                    } else {
                        [void]`$diag.AppendLine("[ERRO] A imagem da pagina nao foi gerada pelo Ghostscript.")
                    }
                }

                Set-Content -Path (Join-Path `$AppFolder "ultimo_texto_extraido.txt") -Value `$diag.ToString() -Encoding UTF8 -ErrorAction SilentlyContinue

                if (`$partes.Count -gt 0) {
                    `$baseName = (`$partes -join '_')
                    `$novoNome = "`${baseName}_`$timestamp.pdf"
                    `$novoPath = Join-Path `$destFolder `$novoNome
                    `$contador = 2
                    while (Test-Path `$novoPath) {
                        `$novoNome = "`${baseName}_`$timestamp (`$contador).pdf"
                        `$novoPath = Join-Path `$destFolder `$novoNome
                        `$contador++
                    }
                    Move-Item -Path `$pdfPath -Destination `$novoPath -Force
                    `$pdfPath = `$novoPath
                }
            }

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
        Set-ConfigValue -Key "SaveFolder" -Value $novaPasta
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

function Set-NamePattern {
    if (-not (Test-Path $APP_FOLDER)) {
        [System.Windows.Forms.MessageBox]::Show("A Impressora PDF Virtual ainda nao foi instalada.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $atual = Get-ConfigValue -Key "NamePatterns"
    $listaAtual = @()
    if ($atual) { $listaAtual = @($atual -split '\|\|\|') }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Configurar Nome Automatico dos PDFs"
    $dlg.Size = New-Object System.Drawing.Size(460, 360)
    $dlg.StartPosition = "CenterScreen"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Digite um rotulo por linha, exatamente como aparece no PDF (ex: transacionador:," + [Environment]::NewLine + "natureza da operacao). O monitor procura cada um e junta o que encontrar" + [Environment]::NewLine + "no nome do arquivo. Deixe tudo em branco para usar somente data/hora."
    $lbl.Size = New-Object System.Drawing.Size(410, 60)
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $dlg.Controls.Add($lbl)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ScrollBars = "Vertical"
    $txt.Size = New-Object System.Drawing.Size(410, 170)
    $txt.Location = New-Object System.Drawing.Point(15, 80)
    $txt.Text = ($listaAtual -join [Environment]::NewLine)
    $dlg.Controls.Add($txt)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Salvar"
    $btnOk.Size = New-Object System.Drawing.Size(120, 35)
    $btnOk.Location = New-Object System.Drawing.Point(15, 265)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 35)
    $btnCancel.Location = New-Object System.Drawing.Point(145, 265)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $novaLista = @($txt.Text -split "`r`n|`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        Set-ConfigValue -Key "NamePatterns" -Value ($novaLista -join '|||')

        if ($novaLista.Count -gt 0) {
            [System.Windows.Forms.MessageBox]::Show(("Rotulos salvos:" + [Environment]::NewLine + ($novaLista -join [Environment]::NewLine)), "Configuracao salva", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Nomeacao automatica desativada. PDFs voltarao a usar somente data/hora.", "Configuracao salva", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    }
}

function Uninstall-Printer {
    Write-Host "================================================================"
    Write-Host "   DESINSTALAR - IMPRESSORA VIRTUAL PDF"
    Write-Host "================================================================"
    Write-Host ""

    if (-not (Test-IsAdmin)) {
        Write-Host "[ERRO] Este processo precisa ser executado como Administrador." -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        return
    }

    $confirmar = [System.Windows.Forms.MessageBox]::Show(
        "Tem certeza que deseja desinstalar a Impressora PDF Virtual?",
        "Confirmar desinstalacao",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirmar -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "[AVISO] Desinstalacao cancelada." -ForegroundColor Yellow
        Read-Host "Pressione Enter para sair"
        return
    }

    Write-Host "[1/5] Parando o monitor em execucao..."
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -like "*monitor_pdf.ps1*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Write-Host "      [OK] Monitor parado" -ForegroundColor Green
    Write-Host ""

    Write-Host "[2/5] Removendo impressora e porta..."
    Remove-Printer -Name $PRINTER_NAME -ErrorAction SilentlyContinue
    Remove-PrinterPort -Name $PORT_FILE -ErrorAction SilentlyContinue
    Write-Host "      [OK] Impressora e porta removidas" -ForegroundColor Green
    Write-Host ""

    Write-Host "[3/5] Removendo inicializacao automatica..."
    Unregister-ScheduledTask -TaskName 'Monitor PDF Virtual' -Confirm:$false -ErrorAction SilentlyContinue
    $wsh = New-Object -ComObject WScript.Shell
    $startupFolder = $wsh.SpecialFolders('AllUsersStartup')
    $shortcutPath = Join-Path $startupFolder 'Monitor PDF Virtual.lnk'
    Remove-Item $shortcutPath -Force -ErrorAction SilentlyContinue
    Write-Host "      [OK] Atalho de inicializacao removido" -ForegroundColor Green
    Write-Host ""

    Write-Host "[4/5] Arquivos e configuracoes..."
    $removerDados = [System.Windows.Forms.MessageBox]::Show(
        "Deseja tambem apagar os arquivos de configuracao e a pasta de destino salva ($APP_FOLDER)?" + [Environment]::NewLine + "Isso remove a memoria da pasta onde os PDFs eram salvos.",
        "Remover configuracoes",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($removerDados -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-Item $APP_FOLDER -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "      [OK] Pasta $APP_FOLDER removida" -ForegroundColor Green
    } else {
        Write-Host "      [OK] Configuracoes mantidas em $APP_FOLDER" -ForegroundColor Yellow
    }
    Write-Host ""

    Write-Host "[5/5] Concluido"
    Write-Host "================================================================"
    Write-Host "   DESINSTALACAO CONCLUIDA!"
    Write-Host "================================================================"
    Write-Host ""
    Read-Host "Pressione Enter para fechar"
}

# --- Execucao direta de uma acao especifica (usado apos auto-elevacao) ---
if ($Action -eq 'Install') { Install-Printer; exit }
if ($Action -eq 'AlterarPasta') { Set-DestinationFolder; exit }
if ($Action -eq 'Uninstall') { Uninstall-Printer; exit }
if ($Action -eq 'NamePattern') { Set-NamePattern; exit }

# --- Interface grafica principal ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Impressora PDF Virtual"
$form.Size = New-Object System.Drawing.Size(420, 365)
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

$btnNomePattern = New-Object System.Windows.Forms.Button
$btnNomePattern.Text = "Configurar Nome Automatico dos PDFs"
$btnNomePattern.Size = New-Object System.Drawing.Size(360, 45)
$btnNomePattern.Location = New-Object System.Drawing.Point(20, 200)
$btnNomePattern.Add_Click({
    $form.Hide()
    Set-NamePattern
    $form.Show()
})
$form.Controls.Add($btnNomePattern)

$btnDesinstalar = New-Object System.Windows.Forms.Button
$btnDesinstalar.Text = "Desinstalar"
$btnDesinstalar.Size = New-Object System.Drawing.Size(360, 45)
$btnDesinstalar.Location = New-Object System.Drawing.Point(20, 255)
$btnDesinstalar.Add_Click({
    $form.Hide()
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Uninstall" -Wait
    $form.Show()
})
$form.Controls.Add($btnDesinstalar)

$btnSair = New-Object System.Windows.Forms.Button
$btnSair.Text = "Sair"
$btnSair.Size = New-Object System.Drawing.Size(360, 30)
$btnSair.Location = New-Object System.Drawing.Point(20, 305)
$btnSair.Add_Click({ $form.Close() })
$form.Controls.Add($btnSair)

[void]$form.ShowDialog()
