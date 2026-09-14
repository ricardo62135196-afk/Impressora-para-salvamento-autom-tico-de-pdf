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
$CONFIG_FILE    = "$APP_FOLDER\config.json"
$LEGACY_CONFIG_FILE = "$APP_FOLDER\config.ini"
$MONITOR_SCRIPT = "$APP_FOLDER\monitor_pdf.ps1"
$PRINTER_NAME   = "Impressora PDF Virtual"
$SAFE_DEFAULT_FOLDER = "$APP_FOLDER\PDFs"

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
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "$env:LOCALAPPDATA\Programs\Tesseract-OCR\tesseract.exe",
        "$env:LOCALAPPDATA\Tesseract-OCR\tesseract.exe"
    )
    foreach ($c in $candidatos) {
        if (Test-Path $c) { return $c }
    }

    # Instalacao feita sem admin cai na pasta do usuario. Como a instalacao da impressora
    # roda elevada (as vezes com conta diferente de quem instalou o Tesseract), procura em
    # TODOS os perfis da maquina, nao so no do processo atual.
    $padroesCoringa = @(
        "C:\Users\*\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        "C:\Users\*\AppData\Local\Tesseract-OCR\tesseract.exe"
    )
    foreach ($padrao in $padroesCoringa) {
        $achado = Get-ChildItem -Path $padrao -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($achado) { return $achado.FullName }
    }

    return $null
}

$SHARED_MATCHING_FUNCTIONS = @'
function Remove-Diacritics {
    param($Texto)
    $normalizado = $Texto.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalizado.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
}

function Get-FlexPattern {
    param($Texto)
    $chars = $Texto.ToCharArray() | ForEach-Object { [RegEx]::Escape([string]$_) }
    return ($chars -join '\s*')
}

function Get-SanitizedName {
    param($Texto)
    $invalidos = [IO.Path]::GetInvalidFileNameChars() -join ''
    $regexInvalidos = "[{0}]" -f [RegEx]::Escape($invalidos)
    $limpo = ($Texto -replace $regexInvalidos, '_').Trim()
    $limpo = $limpo -replace '\s+', ' '
    if ($limpo.Length -gt 60) { $limpo = $limpo.Substring(0, 60).Trim() }
    if ($limpo -eq "") { return $null }
    return $limpo
}
'@
Invoke-Expression $SHARED_MATCHING_FUNCTIONS

function Invoke-Tool {
    param($Exe, [string[]]$ArgList)
    $quoted = $ArgList | ForEach-Object { "`"$_`"" }
    return Start-Process -FilePath $Exe -ArgumentList $quoted -NoNewWindow -Wait -PassThru
}

function Get-DefaultConfig {
    return [PSCustomObject]@{
        SaveFolder    = $SAFE_DEFAULT_FOLDER
        NamePatterns  = @()
        SearchablePdf = $false
        FolderRules   = @()
    }
}

function Import-LegacyIniConfig {
    param($Path)
    $cfg = Get-DefaultConfig
    $linhas = Get-Content $Path -Encoding UTF8 -ErrorAction SilentlyContinue
    foreach ($linha in $linhas) {
        if ($linha -match '^SaveFolder=(.*)$') {
            $cfg.SaveFolder = $Matches[1]
        } elseif ($linha -match '^NamePatterns=(.*)$') {
            $cfg.NamePatterns = @($Matches[1] -split '\|\|\|' | Where-Object { $_ -ne '' })
        } elseif ($linha -match '^SearchablePdf=(.*)$') {
            $cfg.SearchablePdf = ($Matches[1] -eq '1')
        } elseif ($linha -match '^FolderRules=(.*)$') {
            $regras = @()
            foreach ($par in ($Matches[1] -split '\|\|\|')) {
                if ($par -eq '') { continue }
                $pp = $par -split '::', 2
                if ($pp.Count -eq 2) { $regras += [PSCustomObject]@{ Valor = $pp[0]; Pasta = $pp[1] } }
            }
            $cfg.FolderRules = $regras
        }
    }
    return $cfg
}

function Import-AppConfig {
    if (Test-Path $CONFIG_FILE) {
        try {
            $json = Get-Content $CONFIG_FILE -Raw -Encoding UTF8 | ConvertFrom-Json
            $cfg = Get-DefaultConfig
            foreach ($prop in $json.PSObject.Properties) { $cfg.$($prop.Name) = $prop.Value }
            if (-not $cfg.NamePatterns) { $cfg.NamePatterns = @() }
            if (-not $cfg.FolderRules) { $cfg.FolderRules = @() }
            return $cfg
        } catch {
            return Get-DefaultConfig
        }
    }
    if (Test-Path $LEGACY_CONFIG_FILE) {
        $migrado = Import-LegacyIniConfig -Path $LEGACY_CONFIG_FILE
        Save-AppConfig -Config $migrado
        return $migrado
    }
    return Get-DefaultConfig
}

function Save-AppConfig {
    param($Config)
    $Config | ConvertTo-Json -Depth 6 | Set-Content -Path $CONFIG_FILE -Encoding UTF8
}

function Show-PreviewResult {
    param($Texto, $Rotulos)

    $resultDlg = New-Object System.Windows.Forms.Form
    $resultDlg.Text = "Previa do Texto Extraido (OCR)"
    $resultDlg.StartPosition = "CenterScreen"

    $linhasResumo = @()
    if ($Rotulos.Count -gt 0) {
        $textoNorm = Remove-Diacritics -Texto $Texto
        $linhasTexto = $textoNorm -split "`r`n|`n"
        foreach ($rotulo in $Rotulos) {
            $rotuloNorm = Remove-Diacritics -Texto $rotulo
            $flex = Get-FlexPattern -Texto $rotuloNorm
            $valor = $null

            if ($textoNorm -match "(?i)$flex\s*[:\-]?\s*([^\r\n]+)") {
                $candidato = Get-SanitizedName -Texto $Matches[1]
                if ($candidato) { $valor = $candidato }
            }
            if (-not $valor) {
                for ($i = 0; $i -lt $linhasTexto.Count; $i++) {
                    if ($linhasTexto[$i] -match "(?i)$flex") {
                        for ($j = $i + 1; $j -lt $linhasTexto.Count; $j++) {
                            $cand = $linhasTexto[$j].Trim()
                            if ($cand -ne "") { $valor = Get-SanitizedName -Texto $cand; break }
                        }
                        break
                    }
                }
            }
            if ($valor) {
                $linhasResumo += "'$rotulo' -> ENCONTRADO: $valor"
            } else {
                $linhasResumo += "'$rotulo' -> NAO encontrado"
            }
        }
    } else {
        $linhasResumo += "(nenhum rotulo digitado ainda para testar)"
    }

    $lblResumo = New-Object System.Windows.Forms.Label
    $lblResumo.Text = ($linhasResumo -join [Environment]::NewLine)
    $lblResumo.Size = New-Object System.Drawing.Size(520, ([Math]::Max(40, 20 * $linhasResumo.Count)))
    $lblResumo.Location = New-Object System.Drawing.Point(15, 15)
    $resultDlg.Controls.Add($lblResumo)

    $txtResult = New-Object System.Windows.Forms.TextBox
    $txtResult.Multiline = $true
    $txtResult.ReadOnly = $true
    $txtResult.ScrollBars = "Vertical"
    $txtResult.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtResult.Size = New-Object System.Drawing.Size(520, 300)
    $txtResult.Location = New-Object System.Drawing.Point(15, ($lblResumo.Location.Y + $lblResumo.Size.Height + 10))
    $txtResult.Text = $Texto
    $resultDlg.Controls.Add($txtResult)

    $btnFechar = New-Object System.Windows.Forms.Button
    $btnFechar.Text = "Fechar"
    $btnFechar.Size = New-Object System.Drawing.Size(120, 35)
    $btnFechar.Location = New-Object System.Drawing.Point(15, ($txtResult.Location.Y + $txtResult.Size.Height + 10))
    $btnFechar.Add_Click({ $resultDlg.Close() })
    $resultDlg.Controls.Add($btnFechar)

    $resultDlg.ClientSize = New-Object System.Drawing.Size(550, ($btnFechar.Location.Y + $btnFechar.Size.Height + 15))

    [void]$resultDlg.ShowDialog()
}

function Test-ExtracaoTexto {
    param($PadroesTexto)

    $gsPath = Find-Ghostscript
    $tessPath = Find-Tesseract

    if (-not $gsPath) {
        [System.Windows.Forms.MessageBox]::Show("Ghostscript nao encontrado. Instale a impressora primeiro.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }
    if (-not $tessPath) {
        [System.Windows.Forms.MessageBox]::Show("Tesseract nao encontrado. Instale em https://github.com/UB-Mannheim/tesseract/wiki", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = "Arquivos PDF (*.pdf)|*.pdf"
    $ofd.Title = "Selecione um PDF ja gerado para testar"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $pdfTeste = $ofd.FileName
    $tempImg = Join-Path $env:TEMP "preview_ocr.png"
    $tempBase = Join-Path $env:TEMP "preview_ocr"
    $tempTxt = "${tempBase}.txt"
    Remove-Item $tempImg, $tempTxt -Force -ErrorAction SilentlyContinue

    $imgArgs = @("-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=png16m", "-r300", "-dFirstPage=1", "-dLastPage=1", "-sOutputFile=$tempImg", "$pdfTeste")
    Invoke-Tool -Exe $gsPath -ArgList $imgArgs | Out-Null

    if (-not (Test-Path $tempImg)) {
        [System.Windows.Forms.MessageBox]::Show("Nao foi possivel renderizar a pagina do PDF selecionado.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    Invoke-Tool -Exe $tessPath -ArgList @($tempImg, $tempBase, "-l", "por") | Out-Null

    if (-not (Test-Path $tempTxt)) {
        [System.Windows.Forms.MessageBox]::Show("O Tesseract nao gerou texto para esse PDF.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        Remove-Item $tempImg -Force -ErrorAction SilentlyContinue
        return
    }

    $texto = Get-Content $tempTxt -Raw -Encoding UTF8
    Remove-Item $tempImg, $tempTxt -Force -ErrorAction SilentlyContinue

    Show-PreviewResult -Texto $texto -Rotulos $PadroesTexto
}

function Get-FolderRules {
    return @((Import-AppConfig).FolderRules)
}

function Set-FolderRulesConfig {
    param($Regras)
    $cfg = Import-AppConfig
    $cfg.FolderRules = @($Regras | ForEach-Object { [PSCustomObject]@{ Valor = $_.Valor; Pasta = $_.Pasta } })
    Save-AppConfig -Config $cfg
}

function Set-FolderRoutingRules {
    if (-not (Test-Path $APP_FOLDER)) {
        [System.Windows.Forms.MessageBox]::Show("A Impressora PDF Virtual ainda nao foi instalada.", "Erro", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    $regras = [System.Collections.ArrayList]@(Get-FolderRules)
    $pastaEscolhida = @{ Caminho = $null }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Configurar Pastas por Valor Extraido"
    $dlg.ClientSize = New-Object System.Drawing.Size(500, 400)
    $dlg.StartPosition = "CenterScreen"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Quando o PRIMEIRO valor identificado (ex: o transacionador) bater com um" + [Environment]::NewLine + "valor abaixo, o PDF vai para a pasta correspondente em vez da pasta padrao." + [Environment]::NewLine + "Se nao bater com nenhuma regra, usa a pasta padrao normalmente."
    $lbl.Size = New-Object System.Drawing.Size(460, 55)
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $dlg.Controls.Add($lbl)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Size = New-Object System.Drawing.Size(460, 150)
    $lst.Location = New-Object System.Drawing.Point(15, 75)
    foreach ($r in $regras) { [void]$lst.Items.Add("$($r.Valor)  =>  $($r.Pasta)") }
    $dlg.Controls.Add($lst)

    $lblValor = New-Object System.Windows.Forms.Label
    $lblValor.Text = "Valor (exatamente como sai no nome, ex: 1-BRUDA MATRIZ):"
    $lblValor.Size = New-Object System.Drawing.Size(460, 20)
    $lblValor.Location = New-Object System.Drawing.Point(15, 235)
    $dlg.Controls.Add($lblValor)

    $txtValor = New-Object System.Windows.Forms.TextBox
    $txtValor.Size = New-Object System.Drawing.Size(300, 25)
    $txtValor.Location = New-Object System.Drawing.Point(15, 258)
    $dlg.Controls.Add($txtValor)

    $btnEscolherPasta = New-Object System.Windows.Forms.Button
    $btnEscolherPasta.Text = "Escolher Pasta..."
    $btnEscolherPasta.Size = New-Object System.Drawing.Size(160, 27)
    $btnEscolherPasta.Location = New-Object System.Drawing.Point(320, 257)
    $btnEscolherPasta.Add_Click({
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description = "Escolha a pasta para esse valor"
        if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $pastaEscolhida.Caminho = $fbd.SelectedPath
            $btnEscolherPasta.Text = "Pasta selecionada (ok)"
        }
    })
    $dlg.Controls.Add($btnEscolherPasta)

    $btnAdicionar = New-Object System.Windows.Forms.Button
    $btnAdicionar.Text = "Adicionar Regra"
    $btnAdicionar.Size = New-Object System.Drawing.Size(220, 32)
    $btnAdicionar.Location = New-Object System.Drawing.Point(15, 292)
    $btnAdicionar.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtValor.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Digite um valor.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        if (-not $pastaEscolhida.Caminho) {
            [System.Windows.Forms.MessageBox]::Show("Escolha uma pasta.", "Aviso", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
            return
        }
        $novaRegra = [PSCustomObject]@{ Valor = $txtValor.Text.Trim(); Pasta = $pastaEscolhida.Caminho }
        [void]$regras.Add($novaRegra)
        [void]$lst.Items.Add("$($novaRegra.Valor)  =>  $($novaRegra.Pasta)")
        $txtValor.Text = ""
        $pastaEscolhida.Caminho = $null
        $btnEscolherPasta.Text = "Escolher Pasta..."
    })
    $dlg.Controls.Add($btnAdicionar)

    $btnRemover = New-Object System.Windows.Forms.Button
    $btnRemover.Text = "Remover Selecionada"
    $btnRemover.Size = New-Object System.Drawing.Size(220, 32)
    $btnRemover.Location = New-Object System.Drawing.Point(245, 292)
    $btnRemover.Add_Click({
        $idx = $lst.SelectedIndex
        if ($idx -ge 0) {
            $regras.RemoveAt($idx)
            $lst.Items.RemoveAt($idx)
        }
    })
    $dlg.Controls.Add($btnRemover)

    $btnFechar = New-Object System.Windows.Forms.Button
    $btnFechar.Text = "Fechar e Salvar"
    $btnFechar.Size = New-Object System.Drawing.Size(460, 35)
    $btnFechar.Location = New-Object System.Drawing.Point(15, 335)
    $btnFechar.Add_Click({
        Set-FolderRulesConfig -Regras $regras
        $dlg.Close()
    })
    $dlg.Controls.Add($btnFechar)

    [void]$dlg.ShowDialog()
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
    if ($LASTEXITCODE -ne 0) {
        Write-Host "      [AVISO] icacls retornou codigo $LASTEXITCODE - a permissao pode nao ter sido aplicada corretamente." -ForegroundColor Yellow
        Write-Host "               Se 'Alterar Pasta' ou 'Configurar Nome Automatico' falharem depois, rode o instalador novamente." -ForegroundColor Yellow
    }
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
        $DEST_FOLDER = $SAFE_DEFAULT_FOLDER
        New-Item -ItemType Directory -Path $DEST_FOLDER -Force | Out-Null
        Write-Host "      [AVISO] Nenhuma pasta selecionada. Usando pasta padrao." -ForegroundColor Yellow
    }
    Write-Host "      [OK] PDFs serao salvos em: $DEST_FOLDER" -ForegroundColor Green
    $cfgInstalacao = Import-AppConfig
    $cfgInstalacao.SaveFolder = $DEST_FOLDER
    Save-AppConfig -Config $cfgInstalacao
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
    $encontrada = $false
    for ($tentativa = 0; $tentativa -lt 5; $tentativa++) {
        Start-Sleep -Milliseconds 500
        if (Get-Printer -Name $PRINTER_NAME -ErrorAction SilentlyContinue) { $encontrada = $true; break }
    }
    if (-not $encontrada) {
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
    $commonContent = @"
function Remove-Diacritics {
$((Get-Command Remove-Diacritics).Definition)
}

function Get-FlexPattern {
$((Get-Command Get-FlexPattern).Definition)
}

function Get-SanitizedName {
$((Get-Command Get-SanitizedName).Definition)
}

function Invoke-Tool {
$((Get-Command Invoke-Tool).Definition)
}
"@
    Set-Content -Path "$APP_FOLDER\Common.ps1" -Value $commonContent -Encoding UTF8
    Write-Host "      [OK] Funcoes compartilhadas geradas em Common.ps1" -ForegroundColor Green
    $monitorContent = @"
# Monitor da Impressora Virtual PDF
`$ErrorActionPreference = "SilentlyContinue"

`$PortFile    = "$PORT_FILE"
`$SpoolFolder = "$SPOOL_FOLDER"
`$ConfigFile  = "$CONFIG_FILE"
`$GS          = "$GS_PATH"
`$AppFolder   = "$APP_FOLDER"
`$Tesseract   = "$TESS_PATH"

Write-Host "=== MONITOR IMPRESSORA PDF VIRTUAL ===" -ForegroundColor Green
Write-Host "Arquivo de porta: `$PortFile" -ForegroundColor Yellow
Write-Host "Config: `$ConfigFile" -ForegroundColor Yellow

function Write-Log {
    param(`$Mensagem)
    try {
        `$linha = "[`$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] `$Mensagem"
        `$logPath = Join-Path `$AppFolder "monitor.log"
        Add-Content -Path `$logPath -Value `$linha -Encoding UTF8 -ErrorAction SilentlyContinue
        if ((Test-Path `$logPath) -and ((Get-Item `$logPath).Length -gt 500KB)) {
            `$linhasLog = Get-Content `$logPath -Encoding UTF8 -Tail 200
            Set-Content -Path `$logPath -Value `$linhasLog -Encoding UTF8
        }
    } catch { }
}

function Get-Config {
    `$default = [PSCustomObject]@{
        SaveFolder    = "$SAFE_DEFAULT_FOLDER"
        NamePatterns  = @()
        SearchablePdf = `$false
        FolderRules   = @()
    }
    if (-not (Test-Path `$ConfigFile)) { return `$default }
    try {
        `$json = Get-Content `$ConfigFile -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not `$json.SaveFolder) { `$json | Add-Member -NotePropertyName SaveFolder -NotePropertyValue `$default.SaveFolder -Force }
        if (-not `$json.NamePatterns) { `$json | Add-Member -NotePropertyName NamePatterns -NotePropertyValue @() -Force }
        if (`$null -eq `$json.SearchablePdf) { `$json | Add-Member -NotePropertyName SearchablePdf -NotePropertyValue `$false -Force }
        if (-not `$json.FolderRules) { `$json | Add-Member -NotePropertyName FolderRules -NotePropertyValue @() -Force }
        if (-not (Test-Path `$json.SaveFolder)) { New-Item -ItemType Directory -Path `$json.SaveFolder -Force | Out-Null }
        return `$json
    } catch {
        Write-Log "[ERRO] Falha ao ler config.json: `$_"
        return `$default
    }
}

function Wait-FileReady {
    param(`$Path, `$TimeoutSeconds = 60)
    `$inicio = Get-Date
    while (`$true) {
        try {
            `$s = [System.IO.File]::Open(`$Path, 'Open', 'ReadWrite', 'None')
            `$s.Close()
            return `$true
        } catch {
            if (((Get-Date) - `$inicio).TotalSeconds -gt `$TimeoutSeconds) { return `$false }
            Start-Sleep -Milliseconds 100
        }
    }
}

. (Join-Path `$AppFolder "Common.ps1")

Write-Log "Monitor iniciado."
Get-ChildItem -Path `$env:TEMP -Filter "gs_ocr_*" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
Get-ChildItem -Path `$SpoolFolder -Filter "job_*.ps" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
Write-Host "Monitor ativo. Aguardando impressoes..." -ForegroundColor Green

:MonitorLoop while (`$true) {
    if (Test-Path `$PortFile) {
        `$liberado = Wait-FileReady -Path `$PortFile -TimeoutSeconds 60
        if (-not `$liberado) {
            Write-Log "[ERRO] Tempo esgotado aguardando o arquivo de impressao ser liberado. Descartando esta impressao."
            Remove-Item `$PortFile -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 1
            continue MonitorLoop
        }

        # Captura o arquivo imediatamente com um nome unico, liberando o job.ps na hora
        # para a proxima impressao. Sem isso, imprimir rapido em sequencia pode misturar
        # ou perder trabalhos, ja que todos compartilham o mesmo arquivo fixo de porta.
        `$timestamp   = Get-Date -Format "yyyy-MM-dd_HH-mm-ss-fff"
        `$jobFilePath = Join-Path `$SpoolFolder "job_`$timestamp.ps"
        try {
            Move-Item -Path `$PortFile -Destination `$jobFilePath -Force -ErrorAction Stop
        } catch {
            Write-Log "[ERRO] Nao foi possivel capturar o arquivo de impressao (pode ter sido sobrescrito por outra impressao simultanea): `$_"
            Start-Sleep -Seconds 1
            continue MonitorLoop
        }

        `$config     = Get-Config
        `$destFolder = `$config.SaveFolder
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
            "`$jobFilePath"
        )

        try {
            Invoke-Tool -Exe `$GS -ArgList `$gsArgs | Out-Null
        } catch {
            Write-Log "[ERRO] Falha ao converter PDF: `$_"
            Write-Host "[ERRO] Falha ao converter: `$_" -ForegroundColor Red
        }

        if ((Test-Path `$pdfPath) -and ((Get-Item `$pdfPath).Length -gt 1024)) {
            `$padroes = @(`$config.NamePatterns)
            `$pesquisavel = [bool]`$config.SearchablePdf
            `$tessDisponivel = (-not [string]::IsNullOrWhiteSpace(`$Tesseract)) -and (Test-Path `$Tesseract)

            if ((`$padroes.Count -gt 0 -or `$pesquisavel) -and `$tessDisponivel) {
                `$diag = New-Object System.Text.StringBuilder
                [void]`$diag.AppendLine("=== DIAGNOSTICO DE OCR ===")
                [void]`$diag.AppendLine("Data/hora: `$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
                [void]`$diag.AppendLine("PDF processado: `$pdfPath")
                [void]`$diag.AppendLine("Rotulos configurados: `$(if (`$padroes.Count -gt 0) { `$padroes -join ' | ' } else { '(nenhum)' })")
                [void]`$diag.AppendLine("PDF pesquisavel habilitado: `$pesquisavel")
                [void]`$diag.AppendLine("")

                `$partes = @()
                `$textoOriginal = `$null

                if (`$pesquisavel) {
                    `$tiffPath = Join-Path `$env:TEMP "gs_ocr_`$timestamp.tiff"
                    `$tiffArgs = @(
                        "-dNOPAUSE", "-dBATCH", "-dSAFER",
                        "-sDEVICE=tiffgray",
                        "-r300",
                        "-dLastPage=30",
                        "-sOutputFile=`$tiffPath",
                        "`$pdfPath"
                    )
                    try {
                        `$gsTiffResult = Invoke-Tool -Exe `$GS -ArgList `$tiffArgs
                        [void]`$diag.AppendLine("Renderizacao de todas as paginas (TIFF, limite de 30 paginas): codigo de saida `$(`$gsTiffResult.ExitCode)")
                    } catch {
                        [void]`$diag.AppendLine("[ERRO] Falha ao renderizar paginas: `$_")
                    }

                    if (Test-Path `$tiffPath) {
                        `$ocrBase = Join-Path `$env:TEMP "gs_ocr_`$timestamp"
                        `$ocrTxtPath = "`${ocrBase}.txt"
                        `$ocrPdfPath = "`${ocrBase}.pdf"

                        try {
                            `$tessResult = Invoke-Tool -Exe `$Tesseract -ArgList @(`$tiffPath, `$ocrBase, "-l", "por", "txt", "pdf")
                            [void]`$diag.AppendLine("Tesseract (txt+pdf) rodou. Codigo de saida: `$(`$tessResult.ExitCode)")
                        } catch {
                            [void]`$diag.AppendLine("[ERRO] Falha ao executar o Tesseract: `$_")
                        }

                        if (Test-Path `$ocrTxtPath) {
                            `$textoOriginal = Get-Content `$ocrTxtPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                            [void]`$diag.AppendLine("Texto OCR extraido: SIM (`$(if (`$textoOriginal) { `$textoOriginal.Length } else { 0 }) caracteres)")
                            Remove-Item `$ocrTxtPath -Force -ErrorAction SilentlyContinue
                        }

                        if (Test-Path `$ocrPdfPath) {
                            Move-Item -Path `$ocrPdfPath -Destination `$pdfPath -Force
                            [void]`$diag.AppendLine("PDF pesquisavel gerado com sucesso (substituiu o original).")
                        } else {
                            [void]`$diag.AppendLine("[ERRO] O Tesseract nao gerou o PDF pesquisavel. O PDF original (sem texto) foi mantido.")
                        }
                    } else {
                        [void]`$diag.AppendLine("[ERRO] A imagem TIFF de todas as paginas nao foi gerada.")
                    }
                    Remove-Item `$tiffPath -Force -ErrorAction SilentlyContinue
                } elseif (`$padroes.Count -gt 0) {
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
                        `$gsImgResult = Invoke-Tool -Exe `$GS -ArgList `$imgArgs
                        [void]`$diag.AppendLine("Renderizacao da pagina 1 em imagem: codigo de saida `$(`$gsImgResult.ExitCode)")
                    } catch {
                        [void]`$diag.AppendLine("[ERRO] Falha ao renderizar pagina em imagem: `$_")
                    }

                    if (Test-Path `$imgPath) {
                        [void]`$diag.AppendLine("Imagem gerada: SIM (`$((Get-Item `$imgPath).Length) bytes)")
                        `$ocrBase = Join-Path `$env:TEMP "gs_ocr_`$timestamp"
                        `$ocrTxtPath = "`${ocrBase}.txt"

                        try {
                            `$tessResult = Invoke-Tool -Exe `$Tesseract -ArgList @(`$imgPath, `$ocrBase, "-l", "por")
                            [void]`$diag.AppendLine("Tesseract rodou. Codigo de saida: `$(`$tessResult.ExitCode)")
                        } catch {
                            [void]`$diag.AppendLine("[ERRO] Falha ao executar o Tesseract: `$_")
                        }

                        if (Test-Path `$ocrTxtPath) {
                            `$textoOriginal = Get-Content `$ocrTxtPath -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
                            [void]`$diag.AppendLine("Texto OCR extraido: SIM (`$(if (`$textoOriginal) { `$textoOriginal.Length } else { 0 }) caracteres)")
                            Remove-Item `$ocrTxtPath -Force -ErrorAction SilentlyContinue
                        }
                    } else {
                        [void]`$diag.AppendLine("[ERRO] A imagem da pagina nao foi gerada pelo Ghostscript.")
                    }
                    Remove-Item `$imgPath -Force -ErrorAction SilentlyContinue
                }

                if (`$textoOriginal -and `$padroes.Count -gt 0) {
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
                }

                if (`$textoOriginal) {
                    [void]`$diag.AppendLine("")
                    [void]`$diag.AppendLine("--- TEXTO RECONHECIDO PELO OCR (na integra) ---")
                    [void]`$diag.AppendLine(`$textoOriginal)
                }

                if (`$partes.Count -gt 0) {
                    `$pastaFinal = `$destFolder
                    `$regrasPasta = @(`$config.FolderRules)
                    if (`$regrasPasta.Count -gt 0) {
                        `$primeiraParte = `$partes[0]
                        `$regraEncontrada = `$null
                        foreach (`$regra in `$regrasPasta) {
                            if (`$regra.Valor -eq `$primeiraParte) { `$regraEncontrada = `$regra; break }
                        }
                        if (-not `$regraEncontrada) {
                            foreach (`$regra in `$regrasPasta) {
                                if (`$primeiraParte.ToLowerInvariant().Contains(`$regra.Valor.ToLowerInvariant()) -or `$regra.Valor.ToLowerInvariant().Contains(`$primeiraParte.ToLowerInvariant())) {
                                    `$regraEncontrada = `$regra
                                    [void]`$diag.AppendLine("[INFO] Nenhuma regra bateu exatamente, usando correspondencia parcial.")
                                    break
                                }
                            }
                        }
                        if (`$regraEncontrada) {
                            if (Test-Path `$regraEncontrada.Pasta) {
                                `$pastaFinal = `$regraEncontrada.Pasta
                                [void]`$diag.AppendLine("Regra de pasta aplicada: '`$primeiraParte' -> `$pastaFinal")
                            } else {
                                [void]`$diag.AppendLine("[AVISO] Regra de pasta para '`$primeiraParte' aponta para pasta inexistente, usando pasta padrao.")
                            }
                        }
                    }

                    `$baseName = (`$partes -join '_')
                    `$novoNome = "`${baseName}_`$timestamp.pdf"
                    `$novoPath = Join-Path `$pastaFinal `$novoNome
                    `$contador = 2
                    while (Test-Path `$novoPath) {
                        `$novoNome = "`${baseName}_`$timestamp (`$contador).pdf"
                        `$novoPath = Join-Path `$pastaFinal `$novoNome
                        `$contador++
                    }
                    Move-Item -Path `$pdfPath -Destination `$novoPath -Force
                    `$pdfPath = `$novoPath
                }

                Set-Content -Path (Join-Path `$AppFolder "ultimo_texto_extraido.txt") -Value `$diag.ToString() -Encoding UTF8 -ErrorAction SilentlyContinue
            } elseif (`$padroes.Count -gt 0 -or `$pesquisavel) {
                `$diagSimples = "Tesseract nao encontrado - recursos de OCR desativados.`r`nInstale em https://github.com/UB-Mannheim/tesseract/wiki e reinstale a impressora."
                Set-Content -Path (Join-Path `$AppFolder "ultimo_texto_extraido.txt") -Value `$diagSimples -Encoding UTF8 -ErrorAction SilentlyContinue
            }

            Write-Log "PDF salvo com sucesso: `$pdfPath"
            Write-Host "[OK] PDF salvo: `$pdfPath" -ForegroundColor Green
        } else {
            Write-Log "[ERRO] PDF nao foi gerado corretamente (arquivo ausente ou muito pequeno): `$pdfPath"
            Write-Host "[ERRO] PDF nao foi gerado corretamente" -ForegroundColor Red
            Remove-Item `$pdfPath -Force -ErrorAction SilentlyContinue
        }

        Remove-Item `$jobFilePath -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 200
}
"@
    Set-Content -Path $MONITOR_SCRIPT -Value $monitorContent -Encoding UTF8
    Write-Host "      [OK] Monitor criado" -ForegroundColor Green
    Write-Host ""

    Write-Host "[9/9] Configurando inicializacao automatica do monitor..."
    try {
        Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force -ErrorAction Stop
    } catch {
        # Pode "falhar" so porque uma Politica de Grupo (comum em maquinas de escola/empresa)
        # controla isso em um nivel mais especifico - o que importa e o resultado efetivo abaixo.
    }
    $politicaEfetiva = Get-ExecutionPolicy
    $politicasQuePermitemSemBypass = @('Bypass', 'Unrestricted', 'RemoteSigned')
    if ($politicaEfetiva -in $politicasQuePermitemSemBypass) {
        Write-Host "      [OK] Politica de execucao efetiva: $politicaEfetiva (monitor roda sem -Bypass)" -ForegroundColor Green
        $argsMonitor = "-WindowStyle Hidden -File `"$MONITOR_SCRIPT`""
    } else {
        Write-Host "      [AVISO] Politica de execucao efetiva: $politicaEfetiva - mantendo -Bypass no atalho como seguranca." -ForegroundColor Yellow
        $argsMonitor = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$MONITOR_SCRIPT`""
    }

    $wsh = New-Object -ComObject WScript.Shell
    $startupFolder = $wsh.SpecialFolders('AllUsersStartup')
    if (-not (Test-Path $startupFolder)) { New-Item -ItemType Directory -Path $startupFolder -Force | Out-Null }
    $shortcutPath = Join-Path $startupFolder 'Monitor PDF Virtual.lnk'
    $s = $wsh.CreateShortcut($shortcutPath)
    $s.TargetPath = "powershell.exe"
    $s.Arguments = $argsMonitor
    $s.WorkingDirectory = $APP_FOLDER
    $s.Save()
    Write-Host "      [OK] Atalho criado para todos os usuarios" -ForegroundColor Green
    Write-Host ""

    Start-Process powershell.exe -ArgumentList $argsMonitor -WorkingDirectory $APP_FOLDER

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

    $cfgAtual = Import-AppConfig
    $pastaAtual = $cfgAtual.SaveFolder
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
        $cfgAtual.SaveFolder = $novaPasta
        Save-AppConfig -Config $cfgAtual
        Write-Host "[OK] config.json atualizado com sucesso" -ForegroundColor Green
    } catch {
        Write-Host "[ERRO] Nao foi possivel gravar o config.json: $_" -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        return
    }

    Write-Host ""
    Write-Host "Conferindo o que foi gravado:"
    Get-Content $CONFIG_FILE -Encoding UTF8
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

    $cfgNomes = Import-AppConfig
    $listaAtual = @($cfgNomes.NamePatterns)
    $pesquisavelAtual = [bool]$cfgNomes.SearchablePdf

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Configurar Nome Automatico e OCR dos PDFs"
    $dlg.ClientSize = New-Object System.Drawing.Size(460, 430)
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
    $txt.Size = New-Object System.Drawing.Size(410, 150)
    $txt.Location = New-Object System.Drawing.Point(15, 80)
    $txt.Text = ($listaAtual -join [Environment]::NewLine)
    $dlg.Controls.Add($txt)

    $btnTestar = New-Object System.Windows.Forms.Button
    $btnTestar.Text = "Previa do Texto Extraido (testar com um PDF)..."
    $btnTestar.Size = New-Object System.Drawing.Size(410, 35)
    $btnTestar.Location = New-Object System.Drawing.Point(15, 240)
    $btnTestar.Add_Click({
        $rotulosAtuais = @($txt.Text -split "`r`n|`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        Test-ExtracaoTexto -PadroesTexto $rotulosAtuais
    })
    $dlg.Controls.Add($btnTestar)

    $chkPesquisavel = New-Object System.Windows.Forms.CheckBox
    $chkPesquisavel.Text = "Tornar todos os PDFs pesquisaveis (OCR em todas as paginas - mais lento)"
    $chkPesquisavel.Size = New-Object System.Drawing.Size(410, 40)
    $chkPesquisavel.Location = New-Object System.Drawing.Point(15, 285)
    $chkPesquisavel.Checked = $pesquisavelAtual
    $dlg.Controls.Add($chkPesquisavel)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Salvar"
    $btnOk.Size = New-Object System.Drawing.Size(120, 35)
    $btnOk.Location = New-Object System.Drawing.Point(15, 375)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancelar"
    $btnCancel.Size = New-Object System.Drawing.Size(120, 35)
    $btnCancel.Location = New-Object System.Drawing.Point(145, 375)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $novaLista = @($txt.Text -split "`r`n|`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        $cfgNomes.NamePatterns = $novaLista
        $cfgNomes.SearchablePdf = [bool]$chkPesquisavel.Checked
        Save-AppConfig -Config $cfgNomes

        $msg = @()
        if ($novaLista.Count -gt 0) {
            $msg += "Rotulos salvos:"
            $msg += ($novaLista -join [Environment]::NewLine)
        } else {
            $msg += "Nomeacao automatica desativada (somente data/hora)."
        }
        $msg += ""
        $msg += "PDFs pesquisaveis: $(if ($chkPesquisavel.Checked) { 'ATIVADO' } else { 'desativado' })"
        [System.Windows.Forms.MessageBox]::Show(($msg -join [Environment]::NewLine), "Configuracao salva", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
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
$formWidth   = 420
$margemX     = 20
$larguraBtn  = 360
$espacamento = 12
$y = 20

$form = New-Object System.Windows.Forms.Form
$form.Text = "Impressora PDF Virtual"
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$titulo = New-Object System.Windows.Forms.Label
$titulo.Text = "Impressora PDF Virtual"
$titulo.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titulo.AutoSize = $true
$titulo.Location = New-Object System.Drawing.Point($margemX, $y)
$form.Controls.Add($titulo)
$y += $titulo.PreferredHeight + 8

$subtitulo = New-Object System.Windows.Forms.Label
$subtitulo.Text = "Escolha uma opcao:"
$subtitulo.AutoSize = $true
$subtitulo.Location = New-Object System.Drawing.Point($margemX, $y)
$form.Controls.Add($subtitulo)
$y += $subtitulo.PreferredHeight + 20

function Add-PanelButton {
    param($Texto, $Altura, $OnClick)
    $script:btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Texto
    $btn.Size = New-Object System.Drawing.Size($larguraBtn, $Altura)
    $btn.Location = New-Object System.Drawing.Point($margemX, $y)
    $btn.Add_Click($OnClick)
    $form.Controls.Add($btn)
    $script:y += $Altura + $espacamento
    return $btn
}

Add-PanelButton -Texto "Instalar / Reinstalar Impressora" -Altura 45 -OnClick {
    $form.Hide()
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Install" -Wait
    $form.Show()
} | Out-Null

Add-PanelButton -Texto "Alterar Pasta de Destino" -Altura 45 -OnClick {
    $form.Hide()
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action AlterarPasta" -Wait
    $form.Show()
} | Out-Null

Add-PanelButton -Texto "Configurar Nome Automatico e OCR" -Altura 45 -OnClick {
    $form.Hide()
    Set-NamePattern
    $form.Show()
} | Out-Null

Add-PanelButton -Texto "Configurar Pastas por Valor Extraido" -Altura 45 -OnClick {
    $form.Hide()
    Set-FolderRoutingRules
    $form.Show()
} | Out-Null

Add-PanelButton -Texto "Desinstalar" -Altura 45 -OnClick {
    $form.Hide()
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Action Uninstall" -Wait
    $form.Show()
} | Out-Null

$y += 6
Add-PanelButton -Texto "Sair" -Altura 32 -OnClick { $form.Close() } | Out-Null

$y += 12
$form.ClientSize = New-Object System.Drawing.Size($formWidth, $y)

[void]$form.ShowDialog()
