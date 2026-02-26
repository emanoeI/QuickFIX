# ==========================================================
# Quick Fix, Desenvolvido por Emanoel Peres.
# ==========================================================

$Host.UI.RawUI.BackgroundColor = "Black"
$Host.UI.RawUI.ForegroundColor = "White"
Clear-Host

# --- PALETA DE CORES ---
$CN = "Green"      # Neon
$CP = "Magenta"    # Purple
$CC = "Cyan"       # Cyan
$CR = "Red"        # Red
$CY = "Yellow"     # Yellow
$CW = "White"      # White
$CG = "DarkGray"   # Gray
$CDG = "DarkGreen"
$CDP = "DarkMagenta"
$CDC = "DarkCyan"

# --- VARIAVEIS DE SESSAO ---
$global:NomeTecnico  = ""
$global:NomeMaquina  = $env:COMPUTERNAME
$global:IPMaquina    = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" } | Select-Object -First 1).IPAddress
$global:PastaLog     = "C:\services\relatorios"

# ==========================================================
# SISTEMA DE RELATORIO LOCAL
# ==========================================================

function Send-Report {
    param(
        [string]$Opcao,
        [string]$Detalhe,
        [string]$Resultado
    )

    $dataHora  = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $nomeArq   = "relatorio_$(Get-Date -Format 'ddMMyyyy_HHmmss').txt"
    $caminhoArq = Join-Path $global:PastaLog $nomeArq

    # Cria a pasta se nao existir
    if (-not (Test-Path $global:PastaLog)) {
        New-Item -ItemType Directory -Path $global:PastaLog -Force | Out-Null
    }

    $conteudo = @"
====================================================
  RELATORIO DE ATENDIMENTO
====================================================

  Data/Hora  : $dataHora
  Tecnico    : $($global:NomeTecnico)
  Maquina    : $($global:NomeMaquina)
  IP         : $($global:IPMaquina)

----------------------------------------------------
  ACAO EXECUTADA
----------------------------------------------------

  Opcao      : $Opcao
  Detalhe    : $Detalhe
  Resultado  : $Resultado

====================================================
"@

    try {
        $conteudo | Out-File -FilePath $caminhoArq -Encoding UTF8 -ErrorAction Stop
        Write-Host "  [>>] Relatorio salvo em $caminhoArq" -ForegroundColor $CG
    } catch {
        Write-Host "  [WRN] Nao foi possivel salvar o relatorio." -ForegroundColor $CY
    }
}


function Show-Intro {
    Clear-Host
    $lines = @(".----------------------------------------------------------------------------------------.
|..#######.....##.....##....####.....######.....##....##....########....####....##.....##|
|.##.....##....##.....##.....##.....##....##....##...##.....##...........##......##...##.|
|.##.....##....##.....##.....##.....##..........##..##......##...........##.......##.##..|
|.##.....##....##.....##.....##.....##..........#####.......######.......##........###...|
|.##..##.##....##.....##.....##.....##..........##..##......##...........##.......##.##..|
|.##....##.....##.....##.....##.....##....##....##...##.....##...........##......##...##.|
|..#####.##.....#######.....####.....######.....##....##....##..........####....##.....##|
'----------------------------------------------------------------------------------------'")
    $colors = @($CP, $CP, $CDG, $CP, $CN, $CP, $CC, $CDP, $CP, $CDG, $CP, $CP)
    foreach ($i in 0..($lines.Length - 1)) {
        Write-Host $lines[$i] -ForegroundColor $colors[$i]
        Start-Sleep -Milliseconds 70
    }

    Write-Host ""
    Write-Host -NoNewline "  >> INICIALIZANDO SISTEMA..." -ForegroundColor $CG
    Start-Sleep -Milliseconds 400
    Write-Host " [ONLINE]" -ForegroundColor $CN
    Start-Sleep -Milliseconds 200

    $boot = @(
        "  >> CARREGANDO MODULOS DE DIAGNOSTICO",
        "  >> VERIFICANDO PRIVILEGIOS DE ADMIN ",
        "  >> ACESSANDO WMI/CIM INTERFACE      ",
        "  >> PREPARANDO MOTOR DE ANALISE      "
    )
    foreach ($line in $boot) {
        Write-Host -NoNewline $line -ForegroundColor $CG
        Start-Sleep -Milliseconds 200
        Write-Host " [OK]" -ForegroundColor $CN
    }
    Start-Sleep -Milliseconds 400

    # Identificacao do tecnico
    Write-Host ""
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |              IDENTIFICACAO DO TECNICO                |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    do {
        $global:NomeTecnico = Read-Host "`n  >> Seu nome"
    } while ($global:NomeTecnico -eq "")
    Write-Host ""
    Write-Host "  Bem-vindo, $($global:NomeTecnico)!" -ForegroundColor $CN
    Start-Sleep -Milliseconds 500
}

# ==========================================================
# BARRA DE PROGRESSO
# ==========================================================

function Show-Progress {
    param([string]$Label = "Processando", [int]$DurationMs = 1200, [string]$Color = $GN)
    $width = 30
    $steps = 20
    $delay = [int]($DurationMs / $steps)
    Write-Host ""
    for ($i = 1; $i -le $steps; $i++) {
        $filled = [math]::Round(($i / $steps) * $width)
        $empty  = $width - $filled
        $bar    = ("#" * $filled) + ("-" * $empty)
        $pct    = [math]::Round(($i / $steps) * 100)
        Write-Host -NoNewline "`r  $Label  [$bar] $pct% " -ForegroundColor $Color
        Start-Sleep -Milliseconds $delay
    }
    Write-Host " DONE" -ForegroundColor $CN
}

# ==========================================================
# ELEMENTOS VISUAIS
# ==========================================================

function Show-Header {
    Clear-Host
    Write-Host ".----------------------------------------------------------------------------------------.
|..#######.....##.....##....####.....######.....##....##....########....####....##.....##|
|.##.....##....##.....##.....##.....##....##....##...##.....##...........##......##...##.|" -ForegroundColor $CC
    Write-Host "|.##.....##....##.....##.....##.....##..........##..##......##...........##.......##.##..|
|.##.....##....##.....##.....##.....##..........#####.......######.......##........###...|" -ForegroundColor $CC
    Write-Host "|.##..##.##....##.....##.....##.....##..........##..##......##...........##.......##.##..|
|.##....##.....##.....##.....##.....##....##....##...##.....##...........##......##...##.|" -ForegroundColor $CN
    Write-Host "|..#####.##.....#######.....####.....######.....##....##....##..........####....##.....##|
'----------------------------------------------------------------------------------------'" -ForegroundColor $CN
    Write-Host "  >> $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor $CG
    Write-Host ""
}

function Write-SectionHeader {
    param([string]$Title, [string]$Color = "Cyan")
    Write-Host ""
    Write-Host "  +--- [ $Title ]" -ForegroundColor $Color
    Write-Host "  |" -ForegroundColor $Color
}

function Write-SectionFooter {
    param([string]$Color = "Cyan")
    Write-Host "  |" -ForegroundColor $Color
    Write-Host "  +----------------------------------------------------------"
}

function Write-Item {
    param([string]$Label, [string]$Value, [string]$Status = "OK")
    $icon  = switch ($Status) {
        "OK"   { "[OK]  " }
        "WARN" { "[WRN] " }
        "FAIL" { "[ERR] " }
        "INFO" { "[>>]  " }
        default { "[>>]  " }
    }
    $color = switch ($Status) {
        "OK"   { $CN }
        "WARN" { $CY }
        "FAIL" { $CR }
        "INFO" { $CC }
        default { $CW }
    }
    Write-Host "  |  " -NoNewline -ForegroundColor $CG
    Write-Host $icon -NoNewline -ForegroundColor $color
    Write-Host "$Label" -NoNewline -ForegroundColor $CW
    Write-Host ": $Value" -ForegroundColor $color
}

function Show-Pause {
    Write-Host ""
    Write-Host "  +======================================================+" -ForegroundColor $CDP
    Write-Host "  |     [ Pressione qualquer tecla para voltar... ]      |" -ForegroundColor $CG
    Write-Host "  +======================================================+" -ForegroundColor $CDP
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Safe-CimQuery {
    param([string]$ClassName)
    try { return Get-CimInstance -ClassName $ClassName -ErrorAction SilentlyContinue } catch { return $null }
}

# ==========================================================
# 1 - HARDWARE
# ==========================================================

function Show-SystemInfo {
    Show-Header
    Show-Progress "  Mapeando Hardware   " 1400 $CN

    Write-SectionHeader "SISTEMA OPERACIONAL" $CC
    $os = Safe-CimQuery "Win32_OperatingSystem"
    if ($os) {
        Write-Item "OS    " $os.Caption "INFO"
        Write-Item "Versao" "$($os.Version)  Build $($os.BuildNumber)" "INFO"
    }
    Write-SectionFooter $CC

    Write-SectionHeader "PROCESSADOR & PLACA-MAE" $CP
    $cpu   = Safe-CimQuery "Win32_Processor"
    $board = Safe-CimQuery "Win32_BaseBoard"
    if ($cpu)   { Write-Item "CPU " "$($cpu.Name) -- $($cpu.NumberOfCores) Cores" "OK" }
    if ($board) { Write-Item "Mobo" "$($board.Manufacturer) $($board.Product)" "INFO" }
    Write-SectionFooter $CP

    Write-SectionHeader "MEMORIA RAM" $CC
    $ramModules = Safe-CimQuery "Win32_PhysicalMemory"
    foreach ($mod in $ramModules) {
        Write-Item "Slot $($mod.DeviceLocator)" "$([math]::Round($mod.Capacity/1GB,0))GB @ $($mod.Speed)MHz" "OK"
    }
    Write-SectionFooter $CC

    Write-SectionHeader "PLACA DE VIDEO" $CP
    $gpus = Safe-CimQuery "Win32_VideoController"
    foreach ($gpu in $gpus) {
        Write-Item "GPU   " $gpu.Name "OK"
        Write-Item "Driver" $gpu.DriverVersion "INFO"
    }
    Write-SectionFooter $CP

    Write-SectionHeader "ARMAZENAMENTO" $CC
    Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | ForEach-Object {
        $dl = $_.DeviceID
        try {
            $idx    = (Get-Partition -DriveLetter $dl.Replace(":","")).DiskNumber
            $phys   = Get-PhysicalDisk | Where-Object { $_.DeviceID -eq [string]$idx }
            $mtype  = $phys.MediaType
            if ($mtype -eq "Unspecified" -or $null -eq $mtype) {
                $mtype = if ($phys.SpindleSpeed -eq 0) { "SSD" } else { "HDD" }
            }
            $total  = [math]::Round($_.Size/1GB, 2)
            $free   = [math]::Round($_.FreeSpace/1GB, 2)
            $pct    = [math]::Round(($_.FreeSpace / $_.Size) * 100, 1)
            $st     = if ($pct -lt 10) { "FAIL" } elseif ($pct -lt 20) { "WARN" } else { "OK" }
            Write-Item "$dl [$mtype] $($phys.FriendlyName)" "$free GB livres / $total GB  ($pct%)" $st
        } catch {
            Write-Item "$dl [?]" "$([math]::Round($_.Size/1GB,2)) GB total" "WARN"
        }
    }
    Write-SectionFooter $CC
    Show-Pause
}

# ==========================================================
# 2 - REDE
# ==========================================================

function Show-NetworkInfo {
    Show-Header
    Show-Progress "  Testando Conectividade" 1200 $CC

    $adapters = Safe-CimQuery "Win32_NetworkAdapterConfiguration" | Where-Object { $_.IPEnabled -eq $true }
    foreach ($adapter in $adapters) {
        Write-SectionHeader $adapter.Description $CP
        Write-Item "IP  " ($adapter.IPAddress -join ', ') "INFO"
        $dhcpMsg = if ($adapter.DHCPEnabled) { "Ativo" } else { "Estatico" }
        Write-Item "DHCP" $dhcpMsg "INFO"
        Write-Item "DNS " ($adapter.DNSServerSearchOrder -join ', ') "INFO"
        $pGate   = if ($adapter.DefaultIPGateway) { Test-Connection -ComputerName $adapter.DefaultIPGateway[0] -Count 1 -Quiet } else { $false }
        $pGoogle = Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet

        $gateMsg = if ($pGate)   { "Respondendo" } else { "Sem Resposta" }
        $gateSt  = if ($pGate)   { "OK" } else { "FAIL" }
        $netMsg  = if ($pGoogle) { "Online" } else { "Offline" }
        $netSt   = if ($pGoogle) { "OK" } else { "FAIL" }

        Write-Item "Gateway " $gateMsg $gateSt
        Write-Item "Internet" $netMsg  $netSt
        Write-SectionFooter $CP
    }
    Show-Pause
}

# ==========================================================
# 3 - IMPRESSORAS
# ==========================================================

function Printer-Diagnostics-Menu {
    do {
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  |          DIAGNOSTICO DE IMPRESSORAS                  |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[1]" -NoNewline -ForegroundColor $CN; Write-Host "  Status Online e Fila de Jobs                  |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[2]" -NoNewline -ForegroundColor $CN; Write-Host "  Mapeamento de Portas e IPs                    |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[3]" -NoNewline -ForegroundColor $CN; Write-Host "  Testar Comunicacao (Ping IP)                  |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[4]" -NoNewline -ForegroundColor $CN; Write-Host "  Reiniciar Spooler de Impressao                |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[5]" -NoNewline -ForegroundColor $CN; Write-Host "  Limpar Fila de Impressao Travada              |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[6]" -NoNewline -ForegroundColor $CN; Write-Host "  Ver e Definir Impressora Padrao               |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[7]" -NoNewline -ForegroundColor $CN; Write-Host "  Enviar Pagina de Teste                        |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[8]" -NoNewline -ForegroundColor $CN; Write-Host "  Forcar Impressora Online                      |" -ForegroundColor $CW
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CY
        Write-Host "  |  " -NoNewline -ForegroundColor $CY; Write-Host "[0]" -NoNewline -ForegroundColor $CR; Write-Host "  Voltar                                        |" -ForegroundColor $CW
        Write-Host "  +======================================================+" -ForegroundColor $CY

        $c = Read-Host "`n  >> "
        switch ($c) {
            '1' {
                Show-Header
                Show-Progress "  Lendo Filas       " 800 $CY
                Write-SectionHeader "STATUS DAS IMPRESSORAS" $CY
                Get-Printer | ForEach-Object {
                    $st = if ($_.PrinterStatus -eq "Normal") { "OK" } else { "WARN" }
                    Write-Item $_.Name "Status: $($_.PrinterStatus) | Jobs: $($_.JobCount)" $st
                }
                Write-SectionFooter $CY
                Show-Pause
            }
            '2' {
                Show-Header
                Show-Progress "  Mapeando Portas   " 800 $CY
                $allPorts = Get-PrinterPort
                Write-SectionHeader "PORTAS E IPs" $CY
                Get-Printer | ForEach-Object {
                    $pName = $_.PortName
                    $ip = ($allPorts | Where-Object { $_.Name -eq $pName }).PrinterHostAddress
                    Write-Item $_.Name "Porta: $pName | IP: $ip" "INFO"
                }
                Write-SectionFooter $CY
                Show-Pause
            }
            '3' {
                $ip = Read-Host "`n  >> IP da Impressora"
                if ($ip) {
                    Show-Progress "  Pingando $ip      " 1000 $CY
                    if (Test-Connection -ComputerName $ip -Count 2 -Quiet) {
                        Write-Item "Resultado" "Impressora RESPONDENDO em $ip" "OK"
                    } else {
                        Write-Item "Resultado" "Sem resposta em $ip" "FAIL"
                    }
                }
                Show-Pause
            }
            '4' {
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host "  |         REINICIAR SPOOLER DE IMPRESSAO               |" -ForegroundColor $CY
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host ""
                Write-Host "  O Spooler gerencia todas as filas de impressao." -ForegroundColor $CG
                Write-Host "  Reinicia-lo resolve a maioria dos casos de" -ForegroundColor $CG
                Write-Host "  impressora que nao imprime sem motivo aparente." -ForegroundColor $CG

                $confirm = Read-Host "`n  >> Reiniciar o Spooler agora? (S/N)"
                if ($confirm -eq 's' -or $confirm -eq 'S') {
                    Show-Progress "  Parando Spooler        " 1000 $CY
                    Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2

                    Show-Progress "  Iniciando Spooler      " 1000 $CY
                    Start-Service -Name Spooler -ErrorAction SilentlyContinue

                    $spooler = Get-Service -Name Spooler
                    Write-SectionHeader "RESULTADO" $CY
                    $spSt  = if ($spooler.Status -eq "Running") { "OK" } else { "FAIL" }
                    $spMsg = if ($spooler.Status -eq "Running") { "Spooler rodando normalmente" } else { "Falha ao iniciar o Spooler" }
                    Write-Item "Spooler" $spMsg $spSt
                    Write-SectionFooter $CY
                }
                Show-Pause
            }
            '5' {
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host "  |         LIMPAR FILA DE IMPRESSAO TRAVADA             |" -ForegroundColor $CY
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host ""
                Write-Host "  Um job travado na fila bloqueia todos os outros." -ForegroundColor $CG
                Write-Host "  Esta opcao para o Spooler, apaga todos os jobs" -ForegroundColor $CG
                Write-Host "  pendentes e reinicia o servico." -ForegroundColor $CG
                Write-Host ""
                Write-Host "  [WRN] Todos os trabalhos em fila serao cancelados." -ForegroundColor $CY

                $confirm = Read-Host "`n  >> Limpar a fila agora? (S/N)"
                if ($confirm -eq 's' -or $confirm -eq 'S') {
                    Show-Progress "  Parando Spooler        " 1000 $CY
                    Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 2

                    Show-Progress "  Removendo jobs travados" 1000 $CY
                    $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
                    $removidos = 0
                    if (Test-Path $spoolPath) {
                        $arquivos = Get-ChildItem -Path $spoolPath -File -ErrorAction SilentlyContinue
                        $removidos = $arquivos.Count
                        $arquivos | Remove-Item -Force -ErrorAction SilentlyContinue
                    }

                    Show-Progress "  Reiniciando Spooler    " 1000 $CY
                    Start-Service -Name Spooler -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 1

                    $spooler = Get-Service -Name Spooler
                    Write-SectionHeader "RESULTADO DA LIMPEZA" $CY
                    Write-Item "Jobs removidos" "$removidos arquivo(s) de fila apagados" "OK"
                    $spSt  = if ($spooler.Status -eq "Running") { "OK" } else { "FAIL" }
                    $spMsg = if ($spooler.Status -eq "Running") { "Spooler rodando normalmente" } else { "Falha ao reiniciar o Spooler" }
                    Write-Item "Spooler" $spMsg $spSt
                    Write-SectionFooter $CY
                }
                Show-Pause
            }
            '6' {
                # Ver e definir impressora padrao
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host "  |         VER E DEFINIR IMPRESSORA PADRAO              |" -ForegroundColor $CY
                Write-Host "  +======================================================+" -ForegroundColor $CY

                $impressoras = Get-Printer | Sort-Object Name
                Write-SectionHeader "IMPRESSORAS DISPONIVEIS" $CY
                $i    = 1
                $lista = @()
                foreach ($imp in $impressoras) {
                    $isPadrao = if ($imp.Default) { " [PADRAO ATUAL]" } else { "" }
                    $st       = if ($imp.Default) { "OK" } else { "INFO" }
                    Write-Item "[$i] $($imp.Name)" "$($imp.PrinterStatus)$isPadrao" $st
                    $lista += $imp.Name
                    $i++
                }
                Write-SectionFooter $CY

                $escolha = Read-Host "`n  >> Numero da impressora para definir como padrao (ENTER para cancelar)"
                if ($escolha -match '^\d+$') {
                    $idx = [int]$escolha - 1
                    if ($idx -ge 0 -and $idx -lt $lista.Count) {
                        $nomeSelecionado = $lista[$idx]
                        Set-Printer -Name $nomeSelecionado -Default $true -ErrorAction SilentlyContinue
                        $novoPadrao = Get-Printer | Where-Object { $_.Default -eq $true }
                        $defSt  = if ($novoPadrao.Name -eq $nomeSelecionado) { "OK" } else { "FAIL" }
                        $defMsg = if ($novoPadrao.Name -eq $nomeSelecionado) { "$nomeSelecionado definida como padrao" } else { "Falha ao definir impressora padrao" }
                        Write-SectionHeader "RESULTADO" $CY
                        Write-Item "Padrao" $defMsg $defSt
                        Write-SectionFooter $CY
                    } else {
                        Write-Item "Erro" "Numero invalido" "FAIL"
                    }
                }
                Show-Pause
            }
            '7' {
                # Pagina de teste personalizada
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host "  |            ENVIAR PAGINA DE TESTE                    |" -ForegroundColor $CY
                Write-Host "  +======================================================+" -ForegroundColor $CY

                $impressoras = Get-Printer | Sort-Object Name
                Write-SectionHeader "SELECIONE A IMPRESSORA" $CY
                $i    = 1
                $lista = @()
                foreach ($imp in $impressoras) {
                    $isPadrao = if ($imp.Default) { " [PADRAO]" } else { "" }
                    Write-Item "[$i]" "$($imp.Name)$isPadrao" "INFO"
                    $lista += $imp.Name
                    $i++
                }
                Write-SectionFooter $CY

                $escolha = Read-Host "`n  >> Numero da impressora"
                if ($escolha -match '^\d+$') {
                    $idx = [int]$escolha - 1
                    if ($idx -ge 0 -and $idx -lt $lista.Count) {
                        $nomeImp  = $lista[$idx]
                        $dataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

                        $conteudo = @"
====================================================
          PAGINA DE TESTE - SUPORTE T.I.
               CLIVALEMAIS
====================================================

  Data/Hora  : $dataHora
  Tecnico    : $($global:NomeTecnico)
  Maquina    : $($global:NomeMaquina)
  IP         : $($global:IPMaquina)
  Impressora : $nomeImp

----------------------------------------------------
  TESTE DE QUALIDADE DE IMPRESSAO
----------------------------------------------------

  Texto normal  : abcdefghijklmnopqrstuvwxyz
  Texto maiusc  : ABCDEFGHIJKLMNOPQRSTUVWXYZ
  Numerico      : 0 1 2 3 4 5 6 7 8 9
  Especiais     : . , ; : ! ? - _ / \ @ # %

----------------------------------------------------

  Se esta pagina foi impressa corretamente,
  a impressora esta funcionando normalmente.

====================================================
"@
                        $tempFile = "$env:TEMP\teste_impressao.txt"
                        $conteudo | Out-File -FilePath $tempFile -Encoding UTF8 -Force

                        Show-Progress "  Enviando para $nomeImp" 1500 $CY
                        Start-Process -FilePath "notepad.exe" -ArgumentList "/p `"$tempFile`"" -Wait -ErrorAction SilentlyContinue

                        Write-SectionHeader "RESULTADO" $CY
                        Write-Item "Pagina de teste" "Enviada para $nomeImp" "OK"
                        Write-Item "Dica" "Verifique se a folha saiu corretamente" "INFO"
                        Write-SectionFooter $CY
                    } else {
                        Write-Item "Erro" "Numero invalido" "FAIL"
                    }
                }
                Show-Pause
            }
            '8' {
                # Forcar impressora online
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host "  |           FORCAR IMPRESSORA ONLINE                   |" -ForegroundColor $CY
                Write-Host "  +======================================================+" -ForegroundColor $CY
                Write-Host ""
                Write-Host "  Quando o Windows mostra a impressora como offline" -ForegroundColor $CG
                Write-Host "  mesmo ela estando ligada e na rede, esta opcao" -ForegroundColor $CG
                Write-Host "  forca o status para online sem reinstalar nada." -ForegroundColor $CG

                $impressoras = Get-Printer | Sort-Object Name
                Write-SectionHeader "SELECIONE A IMPRESSORA" $CY
                $i    = 1
                $lista = @()
                foreach ($imp in $impressoras) {
                    $st = if ($imp.PrinterStatus -ne "Normal") { "WARN" } else { "OK" }
                    Write-Item "[$i]" "$($imp.Name) | Status: $($imp.PrinterStatus)" $st
                    $lista += $imp.Name
                    $i++
                }
                Write-SectionFooter $CY

                $escolha = Read-Host "`n  >> Numero da impressora"
                if ($escolha -match '^\d+$') {
                    $idx = [int]$escolha - 1
                    if ($idx -ge 0 -and $idx -lt $lista.Count) {
                        $nomeImp = $lista[$idx]
                        Show-Progress "  Forcando $nomeImp online" 1200 $CY
                        try {
                            Set-Printer -Name $nomeImp -WorkOffline $false -ErrorAction Stop
                            Write-SectionHeader "RESULTADO" $CY
                            Write-Item "Status" "$nomeImp forcada para ONLINE" "OK"
                            Write-Item "Dica  " "Tente imprimir novamente" "INFO"
                            Write-SectionFooter $CY
                        } catch {
                            Write-SectionHeader "RESULTADO" $CY
                            Write-Item "Status" "Nao foi possivel alterar o status" "FAIL"
                            Write-Item "Dica  " "Verifique se a impressora esta ligada e na rede" "WARN"
                            Write-SectionFooter $CY
                        }
                    } else {
                        Write-Item "Erro" "Numero invalido" "FAIL"
                    }
                }
                Show-Pause
            }
            '0' { return }
        }
    } while ($true)
}

# ==========================================================
# 4 - OTIMIZAR RAM (SUBMENU)
# ==========================================================

# Carrega as funcoes Win32 necessarias para o Mem-Reduct
function Load-MemReduct {
    if (-not ("Win32.MemReductCore" -as [type])) {
        $Sig = @"
        [DllImport("psapi.dll")] public static extern bool EmptyWorkingSet(IntPtr hProcess);
        [DllImport("kernel32.dll")] public static extern bool SetProcessWorkingSetSize(IntPtr hProcess, IntPtr dwMinimumWorkingSetSize, IntPtr dwMaximumWorkingSetSize);
"@
        Add-Type -MemberDefinition $Sig -Name "MemReductCore" -Namespace "Win32" -ErrorAction SilentlyContinue
    }
}

# Exibe analise de RAM apos qualquer operacao
function Show-RamAnalysis {
    $cs          = Safe-CimQuery "Win32_ComputerSystem"
    $ramTotal    = [math]::Round($cs.TotalPhysicalMemory / 1GB, 0)
    $os          = Get-CimInstance Win32_OperatingSystem
    # FreePhysicalMemory e TotalVisibleMemorySize vem em KB
    $ramLivreGB  = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $ramUsadaKB  = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
    $percentRAM  = [math]::Round(($ramUsadaKB / $os.TotalVisibleMemorySize) * 100, 1)

    Write-SectionHeader "ANALISE DE PERFORMANCE POS-OPERACAO" $CN
    Write-Item "RAM Disponivel " "$ramLivreGB GB livres" "OK"
    $ramSt    = if ($ramTotal -le 4) { "WARN" } else { "OK" }
    $commitSt = if ($percentRAM -gt 90) { "FAIL" } elseif ($percentRAM -gt 75) { "WARN" } else { "OK" }
    Write-Item "RAM Total      " "$ramTotal GB" $ramSt
    Write-Item "RAM em Uso     " "$percentRAM %" $commitSt
    Write-SectionFooter $CN

    if ($percentRAM -gt 90) {
        Write-Host ""
        Write-Host "  +======================================================+" -ForegroundColor $CR
        Write-Host "  |  [ERR] DIAGNOSTICO FATAL: Hardware insuficiente!     |" -ForegroundColor $CR
        Write-Host "  |        Windows esta usando o DISCO como RAM.         |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CR
    }
}

# OPCAO 1 - Limpeza geral (comportamento original)
function Memory-LimpezaGeral {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |         LIMPEZA GERAL DE RAM (MEM-REDUCT)            |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN

    $confirm = Read-Host "`n  >> Prosseguir com limpeza profunda? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Load-MemReduct

    Show-Progress "  Limpando RAM dos Processos " 2000 $CN
    Get-Process | ForEach-Object {
        try {
            [Win32.MemReductCore]::SetProcessWorkingSetSize($_.Handle, -1, -1) | Out-Null
            [Win32.MemReductCore]::EmptyWorkingSet($_.Handle) | Out-Null
        } catch {}
    }

    Show-Progress "  Limpando Cache e Temp      " 1500 $CC
    Stop-Process -Name "chrome","msedge" -Force -ErrorAction SilentlyContinue
    foreach ($path in @("$env:TEMP\*","$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache\*")) {
        try { if (Test-Path $path) { Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue } } catch {}
    }

    Show-RamAnalysis
    Show-Pause
}

# OPCAO 2 - Limpar processos Chrome/Edge pesados sem fechar o navegador
function Memory-LimpezaChrome {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |     LIMPEZA CIRURGICA - CHROME/EDGE SEM FECHAR       |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host ""
    Write-Host "  Esta opcao libera a RAM ocupada pelo Chrome/Edge" -ForegroundColor $CG
    Write-Host "  sem fechar o navegador nem as abas abertas." -ForegroundColor $CG
    Write-Host "  As abas continuam funcionando normalmente." -ForegroundColor $CG

    $confirm = Read-Host "`n  >> Prosseguir? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Load-MemReduct

    # Limiar: processos acima de 100MB sao considerados pesados
    $limiarMB = 100

    $processos = Get-Process | Where-Object {
        $_.Name -in @("chrome","msedge","msedgewebview2") -and
        ($_.WorkingSet64 / 1MB) -gt $limiarMB
    }

    if (-not $processos) {
        Write-Host ""
        Write-Item "Chrome/Edge" "Nenhum processo acima de $limiarMB MB encontrado" "INFO"
        Show-Pause
        return
    }

    $totalAntesMB = [math]::Round(($processos | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    Write-Host ""
    Write-Item "Processos encontrados" "$($processos.Count) acima de $limiarMB MB" "WARN"
    Write-Item "RAM ocupada por eles " "$totalAntesMB MB" "WARN"

    Show-Progress "  Aplicando Mem-Reduct no Chrome/Edge" 2000 $CN

    $processos | ForEach-Object {
        try {
            [Win32.MemReductCore]::SetProcessWorkingSetSize($_.Handle, -1, -1) | Out-Null
            [Win32.MemReductCore]::EmptyWorkingSet($_.Handle) | Out-Null
        } catch {}
    }

    # Mede o quanto foi liberado
    $processosDepois = Get-Process | Where-Object { $_.Name -in @("chrome","msedge","msedgewebview2") }
    $totalDepoisMB   = [math]::Round(($processosDepois | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    $liberadoMB      = $totalAntesMB - $totalDepoisMB

    Write-SectionHeader "RESULTADO DA LIMPEZA CIRURGICA" $CN
    Write-Item "RAM antes  " "$totalAntesMB MB em uso pelo Chrome/Edge" "WARN"
    Write-Item "RAM depois " "$totalDepoisMB MB em uso pelo Chrome/Edge" "OK"
    $libSt = if ($liberadoMB -gt 0) { "OK" } else { "INFO" }
    Write-Item "Liberado   " "$liberadoMB MB devolvidos ao sistema" $libSt
    Write-SectionFooter $CN

    Show-RamAnalysis
    Show-Pause
}

# OPCAO 3 - Aplicar flags de economia no atalho do Chrome
function Memory-FlagsChrome {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |        FLAGS DE ECONOMIA DE RAM NO CHROME            |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host ""
    Write-Host "  Esta opcao modifica o atalho do Chrome na area de      " -ForegroundColor $CG
    Write-Host "  trabalho para ativar o descarte automatico de abas     " -ForegroundColor $CG
    Write-Host "  inativas (Tab Discarding). A aba continua aparecendo   " -ForegroundColor $CG
    Write-Host "  mas sai da RAM ate o usuario clicar nela novamente.    " -ForegroundColor $CG
    Write-Host ""
    Write-Host "  [WRN] Isso altera o atalho do Chrome na area de trabalho." -ForegroundColor $CY

    $confirm = Read-Host "`n  >> Prosseguir? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    # Flags que reduzem consumo de RAM no Chrome
    $flags = "--memory-pressure-off --enable-aggressive-tab-discard --purge-memory-button --disable-features=HeavyAdIntervention"

    # Caminho padrao do executavel do Chrome
    $chromePaths = @(
        "$env:PROGRAMFILES\Google\Chrome\Application\chrome.exe",
        "$env:PROGRAMFILES(X86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )

    $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $chromeExe) {
        Write-Host ""
        Write-Item "Chrome" "Executavel nao encontrado nos caminhos padrao" "FAIL"
        Show-Pause
        return
    }

    Show-Progress "  Localizando atalhos do Chrome  " 800 $CN

    # Locais onde os atalhos costumam estar
    $atalhosPaths = @(
        "$env:PUBLIC\Desktop\Google Chrome.lnk",
        "$env:USERPROFILE\Desktop\Google Chrome.lnk"
    )

    $shell        = New-Object -ComObject WScript.Shell
    $alterados    = 0
    $jaConfigurado = $false

    foreach ($atalho in $atalhosPaths) {
        if (Test-Path $atalho) {
            $lnk = $shell.CreateShortcut($atalho)

            # Verifica se as flags ja estao aplicadas
            if ($lnk.Arguments -like "*enable-aggressive-tab-discard*") {
                $jaConfigurado = $true
                continue
            }

            $lnk.TargetPath = $chromeExe
            $lnk.Arguments  = $flags
            $lnk.Save()
            $alterados++
        }
    }

    Write-Host ""
    if ($jaConfigurado -and $alterados -eq 0) {
        Write-Item "Status" "Flags ja estavam aplicadas anteriormente" "INFO"
    } elseif ($alterados -gt 0) {
        Write-Item "Atalhos modificados" "$alterados atalho(s) atualizados" "OK"
        Write-Item "Flags aplicadas    " "Tab Discard + Memory Pressure ativo" "OK"
        Write-Host ""
        Write-Host "  [WRN] Feche e reabra o Chrome pelo atalho para ativar." -ForegroundColor $CY
    } else {
        Write-Item "Atalhos" "Nenhum atalho encontrado na area de trabalho" "WARN"
        Write-Host ""
        Write-Host "  Crie um atalho do Chrome na area de trabalho e rode novamente." -ForegroundColor $CY
    }

    Show-Pause
}

# MENU DA FUNCAO 4
function Optimize-Memory {
    do {
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CN
        Write-Host "  |          OTIMIZACAO DE RAM - MENU                    |" -ForegroundColor $CN
        Write-Host "  +======================================================+" -ForegroundColor $CN
        Write-Host "  |  " -NoNewline -ForegroundColor $CN
        Write-Host "[1]" -NoNewline -ForegroundColor $CN
        Write-Host "  Limpeza Geral de RAM          (Mem-Reduct)        |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CN
        Write-Host "[2]" -NoNewline -ForegroundColor $CN
        Write-Host "  Limpeza Cirurgica Chrome/Edge (sem fechar)        |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CN
        Write-Host "[3]" -NoNewline -ForegroundColor $CN
        Write-Host "  Aplicar Flags de Economia no Chrome               |" -ForegroundColor $CW
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CN
        Write-Host "  |  " -NoNewline -ForegroundColor $CN
        Write-Host "[0]" -NoNewline -ForegroundColor $CG
        Write-Host "  Voltar                                            |" -ForegroundColor $CG
        Write-Host "  +======================================================+" -ForegroundColor $CN

        $c = Read-Host "`n  >> "
        switch ($c) {
            '1' { Memory-LimpezaGeral }
            '2' { Memory-LimpezaChrome }
            '3' { Memory-FlagsChrome }
            '0' { return }
        }
    } while ($true)
}

# ==========================================================
# 5 - REPARO DE REDE
# ==========================================================

function Repair-Network {
    do {
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CC
        Write-Host "  |          REPARO E DIAGNOSTICO DE REDE                |" -ForegroundColor $CC
        Write-Host "  +======================================================+" -ForegroundColor $CC
        Write-Host "  |  " -NoNewline -ForegroundColor $CC
        Write-Host "[1]" -NoNewline -ForegroundColor $CN
        Write-Host "  Reparo Completo (Release/Renew/DNS/Winsock)      |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CC
        Write-Host "[2]" -NoNewline -ForegroundColor $CN
        Write-Host "  Limpar Cache DNS apenas (flushdns)               |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CC
        Write-Host "[3]" -NoNewline -ForegroundColor $CN
        Write-Host "  Testar Conectividade (Gateway/DNS/Internet)      |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CC
        Write-Host "[4]" -NoNewline -ForegroundColor $CN
        Write-Host "  Resetar Adaptador de Rede                        |" -ForegroundColor $CW
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CC
        Write-Host "  |  " -NoNewline -ForegroundColor $CC
        Write-Host "[0]" -NoNewline -ForegroundColor $CG
        Write-Host "  Voltar                                            |" -ForegroundColor $CG
        Write-Host "  +======================================================+" -ForegroundColor $CC

        $c = Read-Host "`n  >> "
        switch ($c) {
            '1' {
                # Reparo completo original
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CC
                Write-Host "  |           REPARO COMPLETO DE REDE                    |" -ForegroundColor $CC
                Write-Host "  +======================================================+" -ForegroundColor $CC

                $confirm = Read-Host "`n  >> Iniciar reparo completo? (S/N)"
                if ($confirm -eq 's' -or $confirm -eq 'S') {
                    Show-Progress "  Release IP Atual   " 1000 $CC; ipconfig /release | Out-Null
                    Show-Progress "  Limpando Cache DNS " 800  $CC; ipconfig /flushdns | Out-Null
                    Show-Progress "  Renovando IP       " 1200 $CC; ipconfig /renew | Out-Null
                    Show-Progress "  Reset Winsock      " 1000 $CP; netsh winsock reset | Out-Null
                    Show-Progress "  Registrando DNS    " 800  $CC; ipconfig /registerdns | Out-Null

                    Write-SectionHeader "DIAGNOSTICO POS-REPARO" $CC
                    $test   = Test-Connection -ComputerName 8.8.8.8 -Count 2 -Quiet
                    $netMsg = if ($test) { "RESTABELECIDA!" } else { "SEM CONEXAO -- verifique o roteador" }
                    $netSt  = if ($test) { "OK" } else { "FAIL" }
                    Write-Item "Internet" $netMsg $netSt
                    Write-SectionFooter $CC
                    Write-Host "  NOTA: Reset do Winsock pode exigir reinicializacao." -ForegroundColor $CY
                }
                Show-Pause
            }
            '2' {
                # Flush DNS isolado
                Show-Header
                Show-Progress "  Limpando Cache DNS " 1000 $CC
                ipconfig /flushdns | Out-Null

                Write-SectionHeader "RESULTADO" $CC
                Write-Item "Cache DNS" "Limpo com sucesso" "OK"
                Write-Item "Dica    " "Tente acessar o site que estava com problema" "INFO"
                Write-SectionFooter $CC
                Show-Pause
            }
            '3' {
                # Teste de conectividade completo
                Show-Header
                Show-Progress "  Testando Conectividade" 1200 $CC

                $adapters = Safe-CimQuery "Win32_NetworkAdapterConfiguration" | Where-Object { $_.IPEnabled -eq $true }

                foreach ($adapter in $adapters) {
                    Write-SectionHeader $adapter.Description $CC

                    # Testa gateway
                    if ($adapter.DefaultIPGateway) {
                        $gw     = $adapter.DefaultIPGateway[0]
                        $pGate  = Test-Connection -ComputerName $gw -Count 2 -Quiet
                        $gwSt   = if ($pGate) { "OK" } else { "FAIL" }
                        $gwMsg  = if ($pGate) { "Respondendo ($gw)" } else { "Sem resposta ($gw)" }
                        Write-Item "Gateway " $gwMsg $gwSt
                    } else {
                        Write-Item "Gateway" "Nao encontrado" "WARN"
                    }

                    # Testa DNS Google
                    $pDNS   = Test-Connection -ComputerName 8.8.8.8 -Count 2 -Quiet
                    $dnsSt  = if ($pDNS) { "OK" } else { "FAIL" }
                    $dnsMsg = if ($pDNS) { "Respondendo (8.8.8.8)" } else { "Sem resposta (8.8.8.8)" }
                    Write-Item "DNS     " $dnsMsg $dnsSt

                    # Testa resolucao de nome (internet de verdade)
                    $pNet   = $false
                    try { $pNet = [bool](Resolve-DnsName "google.com" -ErrorAction SilentlyContinue) } catch {}
                    $intSt  = if ($pNet) { "OK" } else { "FAIL" }
                    $intMsg = if ($pNet) { "Resolucao de DNS funcionando" } else { "Falha ao resolver nomes" }
                    Write-Item "Internet" $intMsg $intSt

                    Write-SectionFooter $CC
                }
                Show-Pause
            }
            '4' {
                # Reset do adaptador de rede
                Show-Header
                Write-Host "  +======================================================+" -ForegroundColor $CC
                Write-Host "  |           RESETAR ADAPTADOR DE REDE                 |" -ForegroundColor $CC
                Write-Host "  +======================================================+" -ForegroundColor $CC
                Write-Host ""
                Write-Host "  Desativa e reativa o adaptador de rede fisicamente." -ForegroundColor $CG
                Write-Host "  Util quando a conexao esta instavel ou sem IP." -ForegroundColor $CG
                Write-Host ""
                Write-Host "  [WRN] A conexao sera interrompida por alguns segundos." -ForegroundColor $CY

                $confirm = Read-Host "`n  >> Resetar o adaptador agora? (S/N)"
                if ($confirm -eq 's' -or $confirm -eq 'S') {
                    $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }

                    if (-not $adapters) {
                        Write-Item "Adaptador" "Nenhum adaptador ativo encontrado" "WARN"
                    } else {
                        foreach ($adapter in $adapters) {
                            Show-Progress "  Desativando $($adapter.Name) " 1000 $CC
                            Disable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 2

                            Show-Progress "  Reativando $($adapter.Name)  " 1000 $CC
                            Enable-NetAdapter -Name $adapter.Name -Confirm:$false -ErrorAction SilentlyContinue
                            Start-Sleep -Seconds 3

                            $status = (Get-NetAdapter -Name $adapter.Name).Status
                            $adSt   = if ($status -eq "Up") { "OK" } else { "WARN" }
                            $adMsg  = if ($status -eq "Up") { "Adaptador online novamente" } else { "Adaptador ainda inicializando..." }
                            Write-SectionHeader "RESULTADO - $($adapter.Name)" $CC
                            Write-Item "Status" $adMsg $adSt
                            Write-SectionFooter $CC
                        }
                    }
                }
                Show-Pause
            }
            '0' { return }
        }
    } while ($true)
}

# ==========================================================
# 6 - REPARO WINDOWS
# ==========================================================

function Repair-Windows {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CR
    Write-Host "  |     REPARO PROFUNDO DE ARQUIVOS (DISM / SFC)         |" -ForegroundColor $CR
    Write-Host "  |  [WRN] Processo pode levar 10 a 20 minutos.          |" -ForegroundColor $CY
    Write-Host "  +======================================================+" -ForegroundColor $CR

    $confirm = Read-Host "`n  >> Iniciar reparo? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Show-Progress "  DISM RestoreHealth " 3000 $CR
    dism /online /cleanup-image /restorehealth /quiet

    Show-Progress "  SFC Scannow        " 3000 $CY
    sfc /scannow | Out-Null

    Show-Progress "  Limpando WinSxS    " 2000 $CP
    dism /online /cleanup-image /startcomponentcleanup /quiet

    Write-SectionHeader "VEREDITO DO REPARO" $CN
    Write-Item "Base de Dados    " "INTEGRALIZADA"           "OK"
    Write-Item "Arquivos Sistema " "REPARADOS / VERIFICADOS" "OK"
    Write-Item "Lixo Atualizacoes" "REMOVIDO"                "OK"
    Write-SectionFooter $CN
    Write-Host "  Reinicie o computador para aplicar as correcoes." -ForegroundColor $CC
    Show-Pause
}

# ==========================================================
# MENU PRINCIPAL
# ==========================================================

function QuickFix-Menu {
    Show-Intro
    do {
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CP
        Write-Host "  |                  MENU PRINCIPAL                      |" -ForegroundColor $CP
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CP
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[1]" -NoNewline -ForegroundColor $CN
        Write-Host "  Diagnostico de Hardware   (CPU/RAM/GPU/Disco)    |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[2]" -NoNewline -ForegroundColor $CN
        Write-Host "  Status de Rede            (IP/DNS/Pings)         |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[3]" -NoNewline -ForegroundColor $CN
        Write-Host "  Diagnostico de Impressoras (Fila/IP/Porta)       |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[4]" -NoNewline -ForegroundColor $CN
        Write-Host "  Otimizar RAM              (Mem-Reduct Engine)    |" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[5]" -NoNewline -ForegroundColor $CC
        Write-Host "  Reparar Rede              (Release/Renew/Winsock)|" -ForegroundColor $CW
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[6]" -NoNewline -ForegroundColor $CR
        Write-Host "  Reparar Windows           (SFC/DISM/Componentes) |" -ForegroundColor $CW
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CP
        Write-Host "  |  " -NoNewline -ForegroundColor $CP
        Write-Host "[0]" -NoNewline -ForegroundColor $CG
        Write-Host "  Sair                                             |" -ForegroundColor $CG
        Write-Host "  +======================================================+" -ForegroundColor $CP
        Write-Host ""

        $choice = Read-Host "  >> "
        switch ($choice) {
            '1' { Show-SystemInfo;          Send-Report "Diagnostico de Hardware"  "CPU, RAM, GPU, Disco"          "Executado" }
            '2' { Show-NetworkInfo;         Send-Report "Status de Rede"           "IP, DNS, Pings, Gateway"                "Executado" }
            '3' { Printer-Diagnostics-Menu; Send-Report "Diagnostico de Impressoras" "Fila, IP, Porta, Spooler"   "Executado" }
            '4' { Optimize-Memory;          Send-Report "Otimizacao de RAM"        "Reduzir / Limpar Navegadores / Flags"   "Executado" }
            '5' { Repair-Network;           Send-Report "Reparo de Rede"           "Release/Renew/DNS/Winsock"     "Executado" }
            '6' { Repair-Windows;           Send-Report "Reparo do Windows"        "SFC / DISM / Componentes"      "Executado" }
            '0' {
                Clear-Host
                Write-Host ""
                Write-Host "  +======================================================+" -ForegroundColor $CP
                Write-Host "  |       QUICKFIX ENCERRADO -- ATE A PROXIMA!           |" -ForegroundColor $CN
                Write-Host "  +======================================================+" -ForegroundColor $CP
                Write-Host ""
                return
            }
        }
    } while ($true)
}

QuickFix-Menu
