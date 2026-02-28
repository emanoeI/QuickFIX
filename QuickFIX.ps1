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
$global:NomeTecnico    = ""
$global:NomeMaquina    = $env:COMPUTERNAME
$global:IPMaquina      = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" } | Select-Object -First 1).IPAddress
$global:PastaLog       = "C:\services\relatorios"
$global:ArquivoSessao  = ""
$global:InicioSessao   = ""
$KlingoPerfil = "" #  vazio para ser preenchido via menu

# ==========================================================
# SISTEMA DE RELATORIO POR SESSAO
# ==========================================================

function Init-Report {
    $global:InicioSessao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $pastaTecnico = Join-Path $global:PastaLog $global:NomeTecnico

    if (-not (Test-Path $pastaTecnico)) {
        New-Item -ItemType Directory -Path $pastaTecnico -Force | Out-Null
    }

    $nomeArq = "sessao_$(Get-Date -Format 'ddMMyyyy_HHmmss').txt"
    $global:ArquivoSessao = Join-Path $pastaTecnico $nomeArq

    $cabecalho = @"
====================================================
  RELATORIO DE SESSAO - QUICKFIX
====================================================

  Inicio     : $($global:InicioSessao)
  Tecnico    : $($global:NomeTecnico)
  Maquina    : $($global:NomeMaquina)
  IP         : $($global:IPMaquina)

====================================================
  ACOES REALIZADAS
====================================================

"@
    try {
        $cabecalho | Out-File -FilePath $global:ArquivoSessao -Encoding UTF8 -ErrorAction Stop
        Write-Host "  [>>] Log de sessao iniciado em $($global:ArquivoSessao)" -ForegroundColor $CG
    } catch {
        Write-Host "  [WRN] Nao foi possivel criar o arquivo de sessao." -ForegroundColor $CY
    }
}

function Send-Report {
    param(
        [string]$Opcao,
        [string]$Detalhe,
        [string]$Resultado
    )

    if (-not $global:ArquivoSessao -or -not (Test-Path $global:ArquivoSessao)) { return }

    $dataHora = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    $bloco = @"
----------------------------------------------------
  [$dataHora]
  Opcao    : $Opcao
  Detalhe  : $Detalhe
  Resultado: $Resultado

"@
    try {
        $bloco | Out-File -FilePath $global:ArquivoSessao -Encoding UTF8 -Append -ErrorAction Stop
    } catch {
        Write-Host "  [WRN] Nao foi possivel gravar no log da sessao." -ForegroundColor $CY
    }
}

function Close-Report {
    if (-not $global:ArquivoSessao -or -not (Test-Path $global:ArquivoSessao)) { return }

    $fimSessao = Get-Date -Format "dd/MM/yyyy HH:mm:ss"

    $rodape = @"
====================================================
  FIM DA SESSAO
====================================================

  Inicio     : $($global:InicioSessao)
  Encerramento: $fimSessao
  Tecnico    : $($global:NomeTecnico)
  Maquina    : $($global:NomeMaquina)

====================================================
"@
    try {
        $rodape | Out-File -FilePath $global:ArquivoSessao -Encoding UTF8 -Append -ErrorAction Stop
        Write-Host "  [>>] Sessao encerrada. Log salvo em $($global:ArquivoSessao)" -ForegroundColor $CG
    } catch {}
}


function Show-Intro {
    Clear-Host
    # --- LOGO ASCII ---
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

    # --- SEQUENCIA DE BOOT ---
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

    # --- SISTEMA DE LOGIN INTEGRADO ---
    $SenhaUniversal = "suporte" # Alterar senha
    $Autenticado    = $false

    Write-Host ""
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |                IDENTIFICACAO E ACESSO                |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    
    # 1. Identificação do Técnico
    do {
        $global:NomeTecnico = Read-Host "`n  >> Identidade do Tecnico"
    } while ($global:NomeTecnico -eq "")

    # 2. Loop de Senha Invisível
    while (-not $Autenticado) {
        # -AsSecureString oculta a digitação no terminal
        $secureInput = Read-Host "  >> Prezado $($global:NomeTecnico), insira a senha de acesso" -AsSecureString
        
        # Conversão necessária para comparar a senha
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureInput)
        $tentativa = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

        if ($tentativa -eq $SenhaUniversal) {
            $Autenticado = $true
            Write-Host "  [+] Acesso concedido." -ForegroundColor $CN
        } else {
            Write-Host "  [!] Senha incorreta." -ForegroundColor $CR
        }
    }

    # Inicia o log da sessao apenas após o login
    Init-Report

    Write-Host ""
    Write-Host "  Bem vindo, $($global:NomeTecnico)!" -ForegroundColor $CN
    Start-Sleep -Milliseconds 1500
}



function Show-Progress {
    param([string]$Label = "Processando", [int]$DurationMs = 1200, [string]$Color = $CN)
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

    $resumo = @()

    Write-SectionHeader "SISTEMA OPERACIONAL" $CC
    $os = Safe-CimQuery "Win32_OperatingSystem"
    if ($os) {
        Write-Item "OS    " $os.Caption "INFO"
        Write-Item "Versao" "$($os.Version)  Build $($os.BuildNumber)" "INFO"
        $resumo += "OS: $($os.Caption) Build $($os.BuildNumber)"
    }
    Write-SectionFooter $CC

    Write-SectionHeader "PROCESSADOR & PLACA-MAE" $CP
    $cpu   = Safe-CimQuery "Win32_Processor"
    $board = Safe-CimQuery "Win32_BaseBoard"
    if ($cpu)   { Write-Item "CPU " "$($cpu.Name) -- $($cpu.NumberOfCores) Cores" "OK"; $resumo += "CPU: $($cpu.Name)" }
    if ($board) { Write-Item "Mobo" "$($board.Manufacturer) $($board.Product)" "INFO" }
    Write-SectionFooter $CP

    Write-SectionHeader "MEMORIA RAM" $CC
    $ramModules = Safe-CimQuery "Win32_PhysicalMemory"
    foreach ($mod in $ramModules) {
        Write-Item "Slot $($mod.DeviceLocator)" "$([math]::Round($mod.Capacity/1GB,0))GB @ $($mod.Speed)MHz" "OK"
    }
    $totalRam = [math]::Round(($ramModules | Measure-Object -Property Capacity -Sum).Sum / 1GB, 0)
    $resumo += "RAM Total: $($totalRam)GB"
    Write-SectionFooter $CC

    Write-SectionHeader "PLACA DE VIDEO" $CP
    $gpus = Safe-CimQuery "Win32_VideoController"
    foreach ($gpu in $gpus) {
        Write-Item "GPU   " $gpu.Name "OK"
        Write-Item "Driver" $gpu.DriverVersion "INFO"
    }
    Write-SectionFooter $CP

    Write-SectionHeader "ARMAZENAMENTO" $CC
    $discoResumo = @()
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
            $discoResumo += "$dl $mtype $free/$total GB livres ($pct%)"
        } catch {
            Write-Item "$dl [?]" "$([math]::Round($_.Size/1GB,2)) GB total" "WARN"
        }
    }
    $resumo += "Discos: $($discoResumo -join ' | ')"
    Write-SectionFooter $CC

    Send-Report "Diagnostico de Hardware" "CPU, RAM, GPU, Disco" ($resumo -join " | ")
    Show-Pause
}

# ==========================================================
# 2 - REDE
# ==========================================================

function Show-NetworkInfo {
    Show-Header
    Show-Progress "  Testando Conectividade" 1200 $CC

    $resultados = @()
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

        $resultados += "Adaptador: $($adapter.Description) | IP: $($adapter.IPAddress -join ',') | Gateway: $gateMsg | Internet: $netMsg"
    }

    Send-Report "Status de Rede" "IP, DNS, Pings, Gateway" ($resultados -join " || ")
    Show-Pause
}

# ==========================================================
# 3 - IMPRESSORAS
# ==========================================================

# ==========================================================
# 3 - IMPRESSORAS (ATUALIZADO PARA MENU INTERATIVO)
# ==========================================================

# ==========================================================
# 3 - IMPRESSORAS (MENU INTERATIVO + NOVA FOLHA TESTE)
# ==========================================================

function Printer-Diagnostics-Menu {
    $Opcoes = @(
        " [1] Status Online e Fila de Jobs           ",
        " [2] Mapeamento de Portas e IPs             ",
        " [3] Testar Comunicacao (Ping IP)           ",
        " [4] Reiniciar Spooler de Impressao         ",
        " [5] Limpar Fila de Impressao Travada       ",
        " [6] Ver e Definir Impressora Padrao        ",
        " [7] Enviar Pagina de Teste (Diagnostico)   ",
        " [8] Forcar Impressora Online               ",
        " [0] Voltar                                 "
    )
    $Selecao = 0

    while ($true) {
        [System.Console]::CursorVisible = $false
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  |          DIAGNOSTICO DE IMPRESSORAS                  |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CY

        for ($i = 0; $i -lt $Opcoes.Count; $i++) {
            if ($i -eq $Selecao) {
                Write-Host "  > $($Opcoes[$i])" -BackgroundColor $CC -ForegroundColor Black
            } else {
                Write-Host "    $($Opcoes[$i])" -ForegroundColor $CW
            }
        }

        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  Use as setas [CIMA/BAIXO] para navegar e [ENTER] para selecionar." -ForegroundColor $CG

        $Tecla = [System.Console]::ReadKey($true).Key

        if ($Tecla -eq 'UpArrow')   { $Selecao-- }
        if ($Tecla -eq 'DownArrow') { $Selecao++ }

        if ($Selecao -lt 0) { $Selecao = $Opcoes.Count - 1 }
        if ($Selecao -ge $Opcoes.Count) { $Selecao = 0 }

        if ($Tecla -eq 'Enter') {
            [System.Console]::CursorVisible = $true 
            
            switch ($Selecao) {
                0 { # OPCAO 1
                    Show-Header
                    Show-Progress "  Lendo Filas        " 800 $CY
                    Write-SectionHeader "STATUS DAS IMPRESSORAS" $CY
                    $statusResumo = @()
                    Get-Printer | ForEach-Object {
                        $st = if ($_.PrinterStatus -eq "Normal") { "OK" } else { "WARN" }
                        Write-Item $_.Name "Status: $($_.PrinterStatus) | Jobs: $($_.JobCount)" $st
                        $statusResumo += "$($_.Name): $($_.PrinterStatus) | Jobs: $($_.JobCount)"
                    }
                    Write-SectionFooter $CY
                    Send-Report "Impressoras - Status e Fila" "Listagem de impressoras" ($statusResumo -join " | ")
                    Show-Pause
                }
                1 { # OPCAO 2
                    Show-Header
                    Show-Progress "  Mapeando Portas    " 800 $CY
                    $allPorts = Get-PrinterPort
                    Write-SectionHeader "PORTAS E IPs" $CY
                    $portaResumo = @()
                    Get-Printer | ForEach-Object {
                        $pName = $_.PortName
                        $ip = ($allPorts | Where-Object { $_.Name -eq $pName }).PrinterHostAddress
                        Write-Item $_.Name "Porta: $pName | IP: $ip" "INFO"
                        $portaResumo += "$($_.Name) -> Porta: $pName IP: $ip"
                    }
                    Write-SectionFooter $CY
                    Send-Report "Impressoras - Portas e IPs" "Mapeamento de portas" ($portaResumo -join " | ")
                    Show-Pause
                }
                2 { # OPCAO 3
                    $ip = Read-Host "`n  >> IP da Impressora"
                    if ($ip) {
                        Show-Progress "  Pingando $ip      " 1000 $CY
                        $pingOk = Test-Connection -ComputerName $ip -Count 2 -Quiet
                        if ($pingOk) {
                            Write-Item "Resultado" "Impressora RESPONDENDO em $ip" "OK"
                        } else {
                            Write-Item "Resultado" "Sem resposta em $ip" "FAIL"
                        }
                        $pingMsg = if ($pingOk) { "Respondendo" } else { "Sem resposta" }
                        Send-Report "Impressoras - Ping" "IP testado: $ip" "Resultado: $pingMsg"
                    }
                    Show-Pause
                }
                3 { # OPCAO 4
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
                        Send-Report "Impressoras - Reiniciar Spooler" "Reinicio do servico Spooler" $spMsg
                    }
                    Show-Pause
                }
                4 { # OPCAO 5
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
                        Send-Report "Impressoras - Limpar Fila" "Limpeza de jobs travados" "$removidos job(s) removidos | Spooler: $spMsg"
                    }
                    Show-Pause
                }
                5 { # OPCAO 6
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
                            Send-Report "Impressoras - Definir Padrao" "Alteracao de impressora padrao" $defMsg
                        } else {
                            Write-Item "Erro" "Numero invalido" "FAIL"
                        }
                    }
                    Show-Pause
                }
                6 { 6  # OPCAO 7 - ENVIAR PAGINA DE TESTE
                    Show-Header
                    Write-Host "  +======================================================+" -ForegroundColor $CY
                    Write-Host "  |             ENVIAR PAGINA DE TESTE                   |" -ForegroundColor $CY
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
                            Send-Report "Impressoras - Pagina de Teste" "Impressora: $nomeImp" "Pagina de teste enviada"
                        } else {
                            Write-Item "Erro" "Numero invalido" "FAIL"
                        }
                    }
                    Show-Pause
                }
                7 { # OPCAO 8
                    Show-Header
                    Write-Host "  +======================================================+" -ForegroundColor $CY
                    Write-Host "  |            FORCAR IMPRESSORA ONLINE                  |" -ForegroundColor $CY
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
                                Send-Report "Impressoras - Forcar Online" "Impressora: $nomeImp" "Forcada para ONLINE com sucesso"
                            } catch {
                                Write-SectionHeader "RESULTADO" $CY
                                Write-Item "Status" "Nao foi possivel alterar o status" "FAIL"
                                Write-Item "Dica  " "Verifique se a impressora esta ligada e na rede" "WARN"
                                Write-SectionFooter $CY
                                Send-Report "Impressoras - Forcar Online" "Impressora: $nomeImp" "FALHA ao forcar online"
                            }
                        } else {
                            Write-Item "Erro" "Numero invalido" "FAIL"
                        }
                    }
                    Show-Pause
                }
                8 { # VOLTAR (OPCAO 0)
                    return 
                }
            }
        }
    }
}
# ==========================================================
# 4 - OTIMIZAR RAM (SUBMENU)
# ==========================================================

function Load-MemReduct {
    if (-not ("Win32.MemReductCore" -as [type])) {
        $Sig = @"
        [DllImport("psapi.dll")] public static extern bool EmptyWorkingSet(IntPtr hProcess);
        [DllImport("kernel32.dll")] public static extern bool SetProcessWorkingSetSize(IntPtr hProcess, IntPtr dwMinimumWorkingSetSize, IntPtr dwMaximumWorkingSetSize);
"@
        Add-Type -MemberDefinition $Sig -Name "MemReductCore" -Namespace "Win32" -ErrorAction SilentlyContinue
    }
}

function Show-RamAnalysis {
    $cs          = Safe-CimQuery "Win32_ComputerSystem"
    $ramTotal    = [math]::Round($cs.TotalPhysicalMemory / 1GB, 0)
    $os          = Get-CimInstance Win32_OperatingSystem
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

    return $percentRAM
}

function Memory-LimpezaGeral {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |         LIMPEZA GERAL DE RAM (MEM-REDUCT)            |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN

    $confirm = Read-Host "`n  >> Prosseguir com limpeza profunda? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Load-MemReduct

    $osBefore   = Get-CimInstance Win32_OperatingSystem
    $ramAntesMB = [math]::Round(($osBefore.TotalVisibleMemorySize - $osBefore.FreePhysicalMemory) / 1024, 0)

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

    $pct = Show-RamAnalysis
    $osAfter    = Get-CimInstance Win32_OperatingSystem
    $ramDepoisMB = [math]::Round(($osAfter.TotalVisibleMemorySize - $osAfter.FreePhysicalMemory) / 1024, 0)
    $liberadoMB  = $ramAntesMB - $ramDepoisMB

    Send-Report "RAM - Limpeza Geral" "Mem-Reduct + cache Chrome" "Antes: $($ramAntesMB)MB em uso | Depois: $($ramDepoisMB)MB | Liberado: $($liberadoMB)MB | Uso atual: $pct%"
    Show-Pause
}

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

    $limiarMB = 100
    $processos = Get-Process | Where-Object {
        $_.Name -in @("chrome","msedge","msedgewebview2") -and
        ($_.WorkingSet64 / 1MB) -gt $limiarMB
    }

    if (-not $processos) {
        Write-Host ""
        Write-Item "Chrome/Edge" "Nenhum processo acima de $limiarMB MB encontrado" "INFO"
        Send-Report "RAM - Limpeza Chrome/Edge" "Limpeza cirurgica sem fechar" "Nenhum processo acima de $limiarMB MB encontrado"
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

    $processosDepois = Get-Process | Where-Object { $_.Name -in @("chrome","msedge","msedgewebview2") }
    $totalDepoisMB   = [math]::Round(($processosDepois | Measure-Object WorkingSet64 -Sum).Sum / 1MB, 0)
    $liberadoMB      = $totalAntesMB - $totalDepoisMB

    Write-SectionHeader "RESULTADO DA LIMPEZA CIRURGICA" $CN
    Write-Item "RAM antes  " "$totalAntesMB MB em uso pelo Chrome/Edge" "WARN"
    Write-Item "RAM depois " "$totalDepoisMB MB em uso pelo Chrome/Edge" "OK"
    $libSt = if ($liberadoMB -gt 0) { "OK" } else { "INFO" }
    Write-Item "Liberado   " "$liberadoMB MB devolvidos ao sistema" $libSt
    Write-SectionFooter $CN

    Show-RamAnalysis | Out-Null
    Send-Report "RAM - Limpeza Chrome/Edge" "Limpeza cirurgica sem fechar navegador" "Antes: $($totalAntesMB)MB | Depois: $($totalDepoisMB)MB | Liberado: $($liberadoMB)MB"
    Show-Pause
}

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

    $flags = "--memory-pressure-off --enable-aggressive-tab-discard --purge-memory-button --disable-features=HeavyAdIntervention"

    $chromePaths = @(
        "$env:PROGRAMFILES\Google\Chrome\Application\chrome.exe",
        "$env:PROGRAMFILES(X86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )

    $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $chromeExe) {
        Write-Host ""
        Write-Item "Chrome" "Executavel nao encontrado nos caminhos padrao" "FAIL"
        Send-Report "RAM - Flags Chrome" "Aplicar flags de economia" "FALHA: Chrome nao encontrado"
        Show-Pause
        return
    }

    Show-Progress "  Localizando atalhos do Chrome  " 800 $CN

    $atalhosPaths = @(
        "$env:PUBLIC\Desktop\Google Chrome.lnk",
        "$env:USERPROFILE\Desktop\Google Chrome.lnk"
    )

    $shell         = New-Object -ComObject WScript.Shell
    $alterados     = 0
    $jaConfigurado = $false

    foreach ($atalho in $atalhosPaths) {
        if (Test-Path $atalho) {
            $lnk = $shell.CreateShortcut($atalho)
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
    $flagMsg = ""
    if ($jaConfigurado -and $alterados -eq 0) {
        Write-Item "Status" "Flags ja estavam aplicadas anteriormente" "INFO"
        $flagMsg = "Flags ja estavam aplicadas"
    } elseif ($alterados -gt 0) {
        Write-Item "Atalhos modificados" "$alterados atalho(s) atualizados" "OK"
        Write-Item "Flags aplicadas    " "Tab Discard + Memory Pressure ativo" "OK"
        Write-Host ""
        Write-Host "  [WRN] Feche e reabra o Chrome pelo atalho para ativar." -ForegroundColor $CY
        $flagMsg = "$alterados atalho(s) modificados com flags de economia"
    } else {
        Write-Item "Atalhos" "Nenhum atalho encontrado na area de trabalho" "WARN"
        Write-Host ""
        Write-Host "  Crie um atalho do Chrome na area de trabalho e rode novamente." -ForegroundColor $CY
        $flagMsg = "Nenhum atalho do Chrome encontrado"
    }

    Send-Report "RAM - Flags Chrome" "Aplicar flags de economia de RAM" $flagMsg
    Show-Pause
}

function Optimize-Memory {
    $Opcoes = @(
        " [1] Limpeza Geral de RAM          (Mem-Reduct)        ",
        " [2] Limpeza Cirurgica Chrome/Edge (sem fechar)        ",
        " [3] Aplicar Flags de Economia no Chrome               ",
        " [0] Voltar                                            "
    )
    $Selecao = 0
    while ($true) {
        [System.Console]::CursorVisible = $false
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CN
        Write-Host "  |          OTIMIZACAO DE RAM - MENU                    |" -ForegroundColor $CN
        Write-Host "  +======================================================+" -ForegroundColor $CN
        for ($i = 0; $i -lt $Opcoes.Count; $i++) {
            if ($i -eq $Selecao) { Write-Host "  > $($Opcoes[$i])" -BackgroundColor $CC -ForegroundColor Black } 
            else { Write-Host "    $($Opcoes[$i])" -ForegroundColor $CW }
        }
        Write-Host "  +======================================================+" -ForegroundColor $CN
        Write-Host "  Use as setas [CIMA/BAIXO] para navegar e [ENTER] para selecionar." -ForegroundColor $CG

        $Tecla = [System.Console]::ReadKey($true).Key
        if ($Tecla -eq 'UpArrow')   { $Selecao-- }
        if ($Tecla -eq 'DownArrow') { $Selecao++ }
        if ($Selecao -lt 0) { $Selecao = $Opcoes.Count - 1 }
        if ($Selecao -ge $Opcoes.Count) { $Selecao = 0 }

        if ($Tecla -eq 'Enter') {
            [System.Console]::CursorVisible = $true 
            switch ($Selecao) {
                0 { Memory-LimpezaGeral }
                1 { Memory-LimpezaChrome }
                2 { Memory-FlagsChrome }
                3 { return }
            }
        }
    }
}

# ==========================================================
# 5 - REPARO DE REDE
# ==========================================================

function Repair-Network {
    $Opcoes = @(
        " [1] Reparo Completo (Release/Renew/DNS/Winsock)      ",
        " [2] Limpar Cache DNS apenas (flushdns)               ",
        " [3] Testar Conectividade (Gateway/DNS/Internet)      ",
        " [4] Resetar Adaptador de Rede                        ",
        " [0] Voltar                                           "
    )
    $Selecao = 0
    while ($true) {
        [System.Console]::CursorVisible = $false
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CC
        Write-Host "  |          REPARO E DIAGNOSTICO DE REDE                |" -ForegroundColor $CC
        Write-Host "  +======================================================+" -ForegroundColor $CC
        for ($i = 0; $i -lt $Opcoes.Count; $i++) {
            if ($i -eq $Selecao) { Write-Host "  > $($Opcoes[$i])" -BackgroundColor $CC -ForegroundColor Black } 
            else { Write-Host "    $($Opcoes[$i])" -ForegroundColor $CW }
        }
        Write-Host "  +======================================================+" -ForegroundColor $CC
        Write-Host "  Use as setas [CIMA/BAIXO] para navegar e [ENTER] para selecionar." -ForegroundColor $CG

        $Tecla = [System.Console]::ReadKey($true).Key
        if ($Tecla -eq 'UpArrow')   { $Selecao-- }
        if ($Tecla -eq 'DownArrow') { $Selecao++ }
        if ($Selecao -lt 0) { $Selecao = $Opcoes.Count - 1 }
        if ($Selecao -ge $Opcoes.Count) { $Selecao = 0 }

        if ($Tecla -eq 'Enter') {
            [System.Console]::CursorVisible = $true 
            switch ($Selecao) {
                0 { 
                    Show-Header
                    Write-Host "  +======================================================+" -ForegroundColor $CC
                    Write-Host "  |            REPARO COMPLETO DE REDE                   |" -ForegroundColor $CC
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
                        Send-Report "Rede - Reparo Completo" "Release/Renew/DNS/Winsock" "Internet pos-reparo: $netMsg"
                    }
                    Show-Pause
                }
                1 { 
                    Show-Header
                    Show-Progress "  Limpando Cache DNS " 1000 $CC
                    ipconfig /flushdns | Out-Null

                    Write-SectionHeader "RESULTADO" $CC
                    Write-Item "Cache DNS" "Limpo com sucesso" "OK"
                    Write-Item "Dica    " "Tente acessar o site que estava com problema" "INFO"
                    Write-SectionFooter $CC
                    Send-Report "Rede - Flush DNS" "Limpeza de cache DNS" "Cache DNS limpo com sucesso"
                    Show-Pause
                }
                2 { 
                    Show-Header
                    Show-Progress "  Testando Conectividade" 1200 $CC

                    $adapters  = Safe-CimQuery "Win32_NetworkAdapterConfiguration" | Where-Object { $_.IPEnabled -eq $true }
                    $testResumo = @()

                    foreach ($adapter in $adapters) {
                        Write-SectionHeader $adapter.Description $CC

                        if ($adapter.DefaultIPGateway) {
                            $gw     = $adapter.DefaultIPGateway[0]
                            $pGate  = Test-Connection -ComputerName $gw -Count 2 -Quiet
                            $gwSt   = if ($pGate) { "OK" } else { "FAIL" }
                            $gwMsg  = if ($pGate) { "Respondendo ($gw)" } else { "Sem resposta ($gw)" }
                            Write-Item "Gateway " $gwMsg $gwSt
                        } else {
                            Write-Item "Gateway" "Nao encontrado" "WARN"
                            $gwMsg = "Nao encontrado"
                        }

                        $pDNS   = Test-Connection -ComputerName 8.8.8.8 -Count 2 -Quiet
                        $dnsSt  = if ($pDNS) { "OK" } else { "FAIL" }
                        $dnsMsg = if ($pDNS) { "Respondendo (8.8.8.8)" } else { "Sem resposta (8.8.8.8)" }
                        Write-Item "DNS      " $dnsMsg $dnsSt

                        $pNet   = $false
                        try { $pNet = [bool](Resolve-DnsName "google.com" -ErrorAction SilentlyContinue) } catch {}
                        $intSt  = if ($pNet) { "OK" } else { "FAIL" }
                        $intMsg = if ($pNet) { "Resolucao de DNS funcionando" } else { "Falha ao resolver nomes" }
                        Write-Item "Internet" $intMsg $intSt

                        Write-SectionFooter $CC
                        $testResumo += "$($adapter.Description) | GW: $gwMsg | DNS: $dnsMsg | Internet: $intMsg"
                    }

                    Send-Report "Rede - Teste de Conectividade" "Gateway/DNS/Internet" ($testResumo -join " || ")
                    Show-Pause
                }
                3 { 
                    Show-Header
                    Write-Host "  +======================================================+" -ForegroundColor $CC
                    Write-Host "  |            RESETAR ADAPTADOR DE REDE                 |" -ForegroundColor $CC
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
                            Send-Report "Rede - Reset Adaptador" "Desativar/Reativar adaptador" "Nenhum adaptador ativo encontrado"
                        } else {
                            $adResumo = @()
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
                                $adResumo += "$($adapter.Name): $adMsg"
                            }
                            Send-Report "Rede - Reset Adaptador" "Desativar/Reativar adaptador" ($adResumo -join " | ")
                        }
                    }
                    Show-Pause
                }
                4 { return }
            }
        }
    }
}

# ==========================================================
# 6 - REPARO WINDOWS (COM EXIT CODE REAL)
# ==========================================================

function Repair-Windows {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CR
    Write-Host "  |     REPARO PROFUNDO DE ARQUIVOS (DISM / SFC)         |" -ForegroundColor $CR
    Write-Host "  |  [WRN] Processo pode levar 10 a 20 minutos.          |" -ForegroundColor $CY
    Write-Host "  +======================================================+" -ForegroundColor $CR

    $confirm = Read-Host "`n  >> Iniciar reparo? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    # --- DISM RestoreHealth ---
    Show-Progress "  DISM RestoreHealth " 3000 $CR
    dism /online /cleanup-image /restorehealth /quiet
    $dismExitCode = $LASTEXITCODE
    $dismMsg = switch ($dismExitCode) {
        0       { "Concluido com sucesso" }
        87      { "Parametro invalido" }
        2       { "Arquivo nao encontrado" }
        1726    { "Servidor RPC indisponivel" }
        default { "Codigo de saida: $dismExitCode" }
    }
    $dismSt = if ($dismExitCode -eq 0) { "OK" } else { "WARN" }

    # --- SFC Scannow ---
    Show-Progress "  SFC Scannow        " 3000 $CY
    sfc /scannow | Out-Null
    $sfcExitCode = $LASTEXITCODE
    $sfcMsg = switch ($sfcExitCode) {
        0       { "Nenhuma violacao encontrada" }
        1       { "Reparos realizados com sucesso" }
        2       { "Reparos pendentes -- reinicie o PC" }
        3       { "Nao foi possivel reparar alguns arquivos" }
        default { "Codigo de saida: $sfcExitCode" }
    }
    $sfcSt = if ($sfcExitCode -le 1) { "OK" } elseif ($sfcExitCode -eq 2) { "WARN" } else { "FAIL" }

    # --- StartComponentCleanup ---
    Show-Progress "  Limpando WinSxS    " 2000 $CP
    dism /online /cleanup-image /startcomponentcleanup /quiet
    $cleanExitCode = $LASTEXITCODE
    $cleanMsg = if ($cleanExitCode -eq 0) { "Limpeza concluida" } else { "Limpeza com aviso (codigo: $cleanExitCode)" }
    $cleanSt  = if ($cleanExitCode -eq 0) { "OK" } else { "WARN" }

    Write-SectionHeader "VEREDITO DO REPARO" $CN
    Write-Item "DISM RestoreHealth " $dismMsg $dismSt
    Write-Item "SFC Scannow        " $sfcMsg  $sfcSt
    Write-Item "Limpeza WinSxS     " $cleanMsg $cleanSt
    Write-SectionFooter $CN
    Write-Host "  Reinicie o computador para aplicar as correcoes." -ForegroundColor $CC

    Send-Report "Reparo Windows" "DISM/SFC/Componentes" "DISM: $dismMsg | SFC: $sfcMsg | WinSxS: $cleanMsg"
    Show-Pause
}

# ==========================================================
# 7 - LIMPEZA DE PERFIL
# ==========================================================

function Profile-LimpezaTemp {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |        LIMPEZA DE TEMP DO USUARIO                    |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host ""
    Write-Host "  Apaga arquivos temporarios do perfil do usuario logado." -ForegroundColor $CG
    Write-Host "  Nao afeta documentos, fotos ou arquivos pessoais." -ForegroundColor $CG

    # Calcula tamanho antes
    $totalBytes = 0
    $paths = @("$env:TEMP", "$env:LOCALAPPDATA\Temp")
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $totalBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $totalMB = [math]::Round($totalBytes / 1MB, 2)

    Write-Host ""
    Write-Item "Espaco ocupado" "$totalMB MB em arquivos temporarios" "WARN"
    Write-Host ""

    $confirm = Read-Host "  >> Limpar agora? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Show-Progress "  Limpando %TEMP% do usuario " 1200 $CP
    foreach ($p in $paths) {
        if (Test-Path $p) {
            Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Calcula tamanho depois
    $restanteBytes = 0
    foreach ($p in $paths) {
        if (Test-Path $p) {
            $restanteBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $restanteMB = [math]::Round($restanteBytes / 1MB, 2)
    $liberadoMB = [math]::Round($totalMB - $restanteMB, 2)

    Write-SectionHeader "RESULTADO" $CP
    Write-Item "Antes   " "$totalMB MB ocupados" "WARN"
    Write-Item "Liberado" "$liberadoMB MB removidos" "OK"
    Write-Item "Restante" "$restanteMB MB (arquivos em uso pelo sistema)" "INFO"
    Write-SectionFooter $CP

    Send-Report "Limpeza de Perfil - Temp Usuario" "%TEMP% e AppData\Temp" "Antes: $($totalMB)MB | Liberado: $($liberadoMB)MB | Restante: $($restanteMB)MB"
    Show-Pause
}

function Profile-LimpezaSistema {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |        LIMPEZA DE TEMP DO SISTEMA                    |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host ""
    Write-Host "  Apaga arquivos em C:\Windows\Temp." -ForegroundColor $CG
    Write-Host "  Requer privilegios de administrador." -ForegroundColor $CG

    $winTempPath = "$env:SystemRoot\Temp"
    $totalBytes  = 0
    if (Test-Path $winTempPath) {
        $totalBytes = (Get-ChildItem $winTempPath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    }
    $totalMB = [math]::Round($totalBytes / 1MB, 2)

    Write-Host ""
    Write-Item "Espaco ocupado" "$totalMB MB em C:\Windows\Temp" "WARN"
    Write-Host ""

    $confirm = Read-Host "  >> Limpar agora? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Show-Progress "  Limpando C:\Windows\Temp   " 1200 $CP
    if (Test-Path $winTempPath) {
        Get-ChildItem $winTempPath -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    $restanteBytes = 0
    if (Test-Path $winTempPath) {
        $restanteBytes = (Get-ChildItem $winTempPath -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    }
    $restanteMB = [math]::Round($restanteBytes / 1MB, 2)
    $liberadoMB = [math]::Round($totalMB - $restanteMB, 2)

    Write-SectionHeader "RESULTADO" $CP
    Write-Item "Antes   " "$totalMB MB ocupados" "WARN"
    Write-Item "Liberado" "$liberadoMB MB removidos" "OK"
    Write-Item "Restante" "$restanteMB MB (arquivos em uso pelo sistema)" "INFO"
    Write-SectionFooter $CP

    Send-Report "Limpeza de Perfil - Temp Sistema" "C:\Windows\Temp" "Antes: $($totalMB)MB | Liberado: $($liberadoMB)MB | Restante: $($restanteMB)MB"
    Show-Pause
}

function Profile-LimpezaPrefetch {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |        LIMPEZA DE PREFETCH                           |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host ""
    Write-Host "  Apaga arquivos de prefetch do Windows." -ForegroundColor $CG
    Write-Host "  O Windows vai recriar os arquivos necessarios." -ForegroundColor $CG
    Write-Host "  O boot pode ser um pouco mais lento na proxima vez." -ForegroundColor $CG

    $prefetchPath = "$env:SystemRoot\Prefetch"
    $totalBytes   = 0
    if (Test-Path $prefetchPath) {
        $totalBytes = (Get-ChildItem $prefetchPath -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    }
    $totalMB = [math]::Round($totalBytes / 1MB, 2)

    Write-Host ""
    Write-Item "Espaco ocupado" "$totalMB MB em Prefetch" "WARN"
    Write-Host ""

    $confirm = Read-Host "  >> Limpar agora? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Show-Progress "  Limpando Prefetch          " 1000 $CP
    if (Test-Path $prefetchPath) {
        Get-ChildItem $prefetchPath -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    $restanteBytes = 0
    if (Test-Path $prefetchPath) {
        $restanteBytes = (Get-ChildItem $prefetchPath -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
    }
    $restanteMB = [math]::Round($restanteBytes / 1MB, 2)
    $liberadoMB = [math]::Round($totalMB - $restanteMB, 2)

    Write-SectionHeader "RESULTADO" $CP
    Write-Item "Antes   " "$totalMB MB ocupados" "WARN"
    Write-Item "Liberado" "$liberadoMB MB removidos" "OK"
    Write-Item "Restante" "$restanteMB MB" "INFO"
    Write-SectionFooter $CP

    Send-Report "Limpeza de Perfil - Prefetch" "C:\Windows\Prefetch" "Antes: $($totalMB)MB | Liberado: $($liberadoMB)MB | Restante: $($restanteMB)MB"
    Show-Pause
}

function Profile-LimpezaTudo {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host "  |        LIMPEZA COMPLETA DE PERFIL                    |" -ForegroundColor $CP
    Write-Host "  +======================================================+" -ForegroundColor $CP
    Write-Host ""
    Write-Host "  Executa as 3 limpezas em sequencia:" -ForegroundColor $CG
    Write-Host "  - Temp do usuario (%TEMP%)" -ForegroundColor $CG
    Write-Host "  - Temp do sistema (C:\Windows\Temp)" -ForegroundColor $CG
    Write-Host "  - Prefetch" -ForegroundColor $CG

    # Calcula total geral antes
    $allPaths = @(
        "$env:TEMP",
        "$env:LOCALAPPDATA\Temp",
        "$env:SystemRoot\Temp",
        "$env:SystemRoot\Prefetch"
    )
    $totalBytes = 0
    foreach ($p in $allPaths) {
        if (Test-Path $p) {
            $totalBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $totalMB = [math]::Round($totalBytes / 1MB, 2)

    Write-Host ""
    Write-Item "Total a liberar (estimado)" "$totalMB MB" "WARN"
    Write-Host ""

    $confirm = Read-Host "  >> Executar limpeza completa? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Show-Progress "  Limpando Temp Usuario      " 1200 $CP
    foreach ($p in @("$env:TEMP", "$env:LOCALAPPDATA\Temp")) {
        if (Test-Path $p) {
            Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    Show-Progress "  Limpando Temp Sistema      " 1200 $CP
    if (Test-Path "$env:SystemRoot\Temp") {
        Get-ChildItem "$env:SystemRoot\Temp" -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    Show-Progress "  Limpando Prefetch          " 1000 $CP
    if (Test-Path "$env:SystemRoot\Prefetch") {
        Get-ChildItem "$env:SystemRoot\Prefetch" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    # Calcula total restante
    $restanteBytes = 0
    foreach ($p in $allPaths) {
        if (Test-Path $p) {
            $restanteBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $restanteMB = [math]::Round($restanteBytes / 1MB, 2)
    $liberadoMB = [math]::Round($totalMB - $restanteMB, 2)

    Write-SectionHeader "RESULTADO DA LIMPEZA COMPLETA" $CP
    Write-Item "Antes   " "$totalMB MB ocupados" "WARN"
    Write-Item "Liberado" "$liberadoMB MB removidos" "OK"
    Write-Item "Restante" "$restanteMB MB (arquivos em uso)" "INFO"
    Write-SectionFooter $CP

    Send-Report "Limpeza de Perfil - Completa" "Temp Usuario + Temp Sistema + Prefetch" "Antes: $($totalMB)MB | Liberado: $($liberadoMB)MB | Restante: $($restanteMB)MB"
    Show-Pause
}

function Clean-Profile {
    $Opcoes = @(
        " [1] Limpar Temp do Usuario  (%TEMP%)                  ",
        " [2] Limpar Temp do Sistema  (C:\Windows\Temp)         ",
        " [3] Limpar Prefetch                                   ",
        " [4] Limpeza Completa        (Tudo de uma vez)         ",
        " [0] Voltar                                            "
    )
    $Selecao = 0
    while ($true) {
        [System.Console]::CursorVisible = $false
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CP
        Write-Host "  |          LIMPEZA DE PERFIL - MENU                    |" -ForegroundColor $CP
        Write-Host "  +======================================================+" -ForegroundColor $CP
        for ($i = 0; $i -lt $Opcoes.Count; $i++) {
            if ($i -eq $Selecao) { Write-Host "  > $($Opcoes[$i])" -BackgroundColor $CC -ForegroundColor Black } 
            else { Write-Host "    $($Opcoes[$i])" -ForegroundColor $CW }
        }
        Write-Host "  +======================================================+" -ForegroundColor $CP
        Write-Host "  Use as setas [CIMA/BAIXO] para navegar e [ENTER] para selecionar." -ForegroundColor $CG

        $Tecla = [System.Console]::ReadKey($true).Key
        if ($Tecla -eq 'UpArrow')   { $Selecao-- }
        if ($Tecla -eq 'DownArrow') { $Selecao++ }
        if ($Selecao -lt 0) { $Selecao = $Opcoes.Count - 1 }
        if ($Selecao -ge $Opcoes.Count) { $Selecao = 0 }

        if ($Tecla -eq 'Enter') {
            [System.Console]::CursorVisible = $true 
            switch ($Selecao) {
                0 { Profile-LimpezaTemp }
                1 { Profile-LimpezaSistema }
                2 { Profile-LimpezaPrefetch }
                3 { Profile-LimpezaTudo }
                4 { return }
            }
        }
    }
}

# ==========================================================
# 8 - REINSTALAR KLINGO
# ==========================================================

# ----------------------------------------------------------
# CONFIGURACAO DO KLINGO
#
# $KlingoURL     --> URL do Klingo da clinica
# $KlingoPerfil  --> Perfil do Chrome onde o Klingo esta instalado
#                    Para descobrir: rodar o comando abaixo no PowerShell:
#                    Get-ChildItem "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
#                    ForEach-Object { Get-ItemProperty $_.PSPath } |
#                    Where-Object { $_.DisplayName -like "*klingo*" } |
#                    Select-Object -ExpandProperty UninstallString
# $KlingoAppId   --> ID do app PWA do Klingo no Chrome
#                    Obtido no UninstallString acima (--uninstall-app-id=XXXXX)
# $KlingoNome    --> Nome exato como aparece em chrome://apps
# ----------------------------------------------------------
$KlingoURL    = "https://clivale.klingo.app/#/"
$KlingoPerfil = "Profile 3"
$KlingoAppId  = "kddmpnlcfioihmciejceonmlnjpikiic"
$KlingoNome   = "klingo"

function Klingo-DesinstalarApp {
    Write-Host ""
    Write-Host "  Procurando instalacao existente do Klingo..." -ForegroundColor $CG

    # Busca no registro pelo nome do app
    $registryCaminhos = @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $appEncontrado = $null
    foreach ($caminho in $registryCaminhos) {
        if (Test-Path $caminho) {
            $apps = Get-ChildItem $caminho -ErrorAction SilentlyContinue
            foreach ($app in $apps) {
                $props = Get-ItemProperty $app.PSPath -ErrorAction SilentlyContinue
                if ($props.DisplayName -like "*$KlingoNome*") {
                    $appEncontrado = $props
                    break
                }
            }
        }
        if ($appEncontrado) { break }
    }

    if ($appEncontrado) {
        Write-Item "Klingo encontrado" $appEncontrado.DisplayName "WARN"
        Write-Item "Desinstalando    " "Aguarde..." "INFO"

        try {
            Show-Progress "  Desinstalando Klingo       " 2000 $CY

            # Localiza o Chrome
            $chromePaths = @(
                "$env:PROGRAMFILES\Google\Chrome\Application\chrome.exe",
                "$env:PROGRAMFILES(X86)\Google\Chrome\Application\chrome.exe",
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
            )
            $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

            # Chama o uninstaller exato conforme registrado pelo Chrome
            $argDesinstalar = "--profile-directory=`"$KlingoPerfil`" --uninstall-app-id=$KlingoAppId"
            Start-Process -FilePath $chromeExe -ArgumentList $argDesinstalar -Wait -WindowStyle Hidden -ErrorAction Stop
            Start-Sleep -Seconds 3

            Write-Item "Desinstalacao" "Klingo removido com sucesso" "OK"
            return $true
        } catch {
            Write-Item "Desinstalacao" "Falha ao remover: $_" "WARN"
            Write-Item "Dica         " "Continuando com reinstalacao mesmo assim..." "INFO"
            return $false
        }
    } else {
        Write-Item "Klingo" "Nenhuma instalacao anterior encontrada" "INFO"
        return $false
    }
}

function Klingo-InstalarApp {
    # Localiza o Chrome
    $chromePaths = @(
        "$env:PROGRAMFILES\Google\Chrome\Application\chrome.exe",
        "$env:PROGRAMFILES(X86)\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )
    $chromeExe = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $chromeExe) {
        Write-Item "Chrome" "Executavel nao encontrado" "FAIL"
        Write-Item "Dica  " "Verifique se o Google Chrome esta instalado" "WARN"
        return $false
    }

    Write-Item "Chrome" "Encontrado em $chromeExe" "OK"
    Write-Item "Perfil" "Usando $KlingoPerfil" "INFO"
    Show-Progress "  Abrindo Chrome no Klingo   " 1500 $CN

    try {
        # Abre o Chrome direto na URL do Klingo no perfil correto
        $argInstalar = "--profile-directory=`"$KlingoPerfil`" `"$KlingoURL`""
        Start-Process -FilePath $chromeExe -ArgumentList $argInstalar -ErrorAction Stop
        Start-Sleep -Seconds 3

        # Guia visual para o tecnico concluir a instalacao manualmente
        Write-Host ""
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  |         SIGA OS PASSOS ABAIXO NO CHROME              |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host "  |                                                      |" -ForegroundColor $CY
        Write-Host "  |  1. No Chrome aberto, clique nos 3 pontinhos (...)   |" -ForegroundColor $CW
        Write-Host "  |     no canto superior direito                        |" -ForegroundColor $CW
        Write-Host "  |                                                      |" -ForegroundColor $CY
        Write-Host "  |  2. Clique em:                                       |" -ForegroundColor $CW
        Write-Host "  |     >> Transmitir, Guardar e Compartilhar            |" -ForegroundColor $CN
        Write-Host "  |                                                      |" -ForegroundColor $CY
        Write-Host "  |  3. Clique em:                                       |" -ForegroundColor $CW
        Write-Host "  |     >> Salvar como App                               |" -ForegroundColor $CN
        Write-Host "  |                                                      |" -ForegroundColor $CY
        Write-Host "  |  4. Confirme clicando em:                            |" -ForegroundColor $CW
        Write-Host "  |     >> Instalar                                      |" -ForegroundColor $CN
        Write-Host "  |                                                      |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CY
        Write-Host ""

        return $true
    } catch {
        Write-Item "Instalacao" "Falha ao abrir Chrome: $_" "FAIL"
        return $false
    }
}

function Reinstall-Klingo {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |          REINSTALAR KLINGO COMO APP                  |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host ""
    Write-Host "  Este modulo remove a instalacao atual do Klingo" -ForegroundColor $CG
    Write-Host "  e reinstala como app do Chrome de forma limpa." -ForegroundColor $CG
    Write-Host ""
    Write-Host "  URL configurada : $KlingoURL" -ForegroundColor $CC
    Write-Host "  App configurado : $KlingoNome" -ForegroundColor $CC
    Write-Host ""
    Write-Host "  [WRN] O Chrome sera aberto durante o processo." -ForegroundColor $CY

    $confirm = Read-Host "`n  >> Iniciar reinstalacao do Klingo? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    Write-SectionHeader "ETAPA 1 - DESINSTALACAO" $CY
    $desinstalou = Klingo-DesinstalarApp
    Write-SectionFooter $CY

    Write-SectionHeader "ETAPA 2 - INSTALACAO" $CN
    $instalou = Klingo-InstalarApp

    if ($instalou) {
        Write-Item "Chrome" "Aberto direto no Klingo" "OK"
        Write-Item "Acao  " "Siga o guia acima para concluir a instalacao" "WARN"
    } else {
        Write-Item "Chrome" "Nao foi possivel abrir o Chrome" "FAIL"
        Write-Item "Dica  " "Abra o Chrome manualmente e acesse $KlingoURL" "INFO"
    }
    Write-SectionFooter $CN

    $desinstMsg = if ($desinstalou) { "versao anterior removida" } else { "sem instalacao anterior" }
    $instalMsg  = if ($instalou)    { "Chrome aberto no Klingo, aguardando instalacao manual" } else { "FALHA ao abrir Chrome" }
    Send-Report "Klingo - Reinstalacao" "URL: $KlingoURL" "Desinstalacao: $desinstMsg | Instalacao: $instalMsg"

    Show-Pause
}

function Klingo-LimparCache {
    Show-Header
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host "  |          LIMPAR CACHE DO KLINGO                      |" -ForegroundColor $CN
    Write-Host "  +======================================================+" -ForegroundColor $CN
    Write-Host ""
    Write-Host "  Remove cache e dados de armazenamento local do Klingo." -ForegroundColor $CG
    Write-Host "  Resolve lentidao e app que nao abre corretamente." -ForegroundColor $CG
    Write-Host ""
    Write-Host "  [WRN] O Chrome sera fechado durante o processo." -ForegroundColor $CY
    Write-Host "  [WRN] O usuario precisara fazer login no Klingo novamente." -ForegroundColor $CY

    $confirm = Read-Host "`n  >> Limpar cache do Klingo agora? (S/N)"
    if ($confirm -ne 's' -and $confirm -ne 'S') { return }

    # Caminhos do cache e dados do PWA no perfil correto
    $perfilBase = "$env:LOCALAPPDATA\Google\Chrome\User Data\$KlingoPerfil"
    $caminhos   = @(
        # Cache geral do perfil
        "$perfilBase\Cache",
        "$perfilBase\Code Cache",
        # Cache especifico do app (identificado pelo AppId)
        "$perfilBase\Web Applications\_crx_$KlingoAppId\Default\Cache",
        "$perfilBase\Web Applications\_crx_$KlingoAppId\Default\Code Cache",
        # Dados de armazenamento local e IndexedDB do app
        "$perfilBase\Local Storage\leveldb",
        "$perfilBase\IndexedDB\chrome-extension_${KlingoAppId}_0.indexeddb.leveldb",
        "$perfilBase\Service Worker\CacheStorage"
    )

    # Calcula espaco ocupado antes
    $totalBytes = 0
    foreach ($p in $caminhos) {
        if (Test-Path $p) {
            $totalBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $totalMB = [math]::Round($totalBytes / 1MB, 2)
    Write-Host ""
    Write-Item "Cache encontrado" "$totalMB MB" "WARN"

    # Fecha o Chrome antes de limpar
    Show-Progress "  Fechando Chrome            " 1000 $CN
    Get-Process -Name "chrome" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Limpa os caminhos
    Show-Progress "  Limpando Cache do Klingo   " 1500 $CN
    foreach ($p in $caminhos) {
        if (Test-Path $p) {
            Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Calcula espaco restante
    $restanteBytes = 0
    foreach ($p in $caminhos) {
        if (Test-Path $p) {
            $restanteBytes += (Get-ChildItem $p -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        }
    }
    $restanteMB = [math]::Round($restanteBytes / 1MB, 2)
    $liberadoMB = [math]::Round($totalMB - $restanteMB, 2)

    Write-SectionHeader "RESULTADO" $CN
    Write-Item "Antes   " "$totalMB MB de cache" "WARN"
    Write-Item "Liberado" "$liberadoMB MB removidos" "OK"
    Write-Item "Restante" "$restanteMB MB (arquivos em uso)" "INFO"
    Write-Item "Dica    " "Abra o Klingo e faca login novamente" "INFO"
    Write-SectionFooter $CN

    Send-Report "Klingo - Limpeza de Cache" "Cache + LocalStorage + IndexedDB + ServiceWorker" "Antes: $($totalMB)MB | Liberado: $($liberadoMB)MB | Restante: $($restanteMB)MB"
    Show-Pause
}

function Klingo-Menu {
    # Se ainda não escolheu o perfil nesta sessão, pergunta agora
    if ($global:KlingoPerfil -eq "") {
        $global:KlingoPerfil = Get-Chrome-Profile-List
    }

    do {
        Show-Header
        Write-Host "  +======================================================+" -ForegroundColor $CN
        Write-Host "  |                KLINGO - MENU                         |" -ForegroundColor $CN
        Write-Host "  |  Perfil Atual: $($global:KlingoPerfil.PadRight(34))  |" -ForegroundColor $CY
        Write-Host "  +======================================================+" -ForegroundColor $CN
        
        Write-Host "  |  [1] Reinstalar Klingo (Neste Perfil)              |" -ForegroundColor $CW
        Write-Host "  |  [2] Limpar Cache      (Neste Perfil)              |" -ForegroundColor $CW
        Write-Host "  |  [3] Trocar Perfil Selecionado                     |" -ForegroundColor $CC
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CN
        Write-Host "  |  [0] Voltar                                        |" -ForegroundColor $CG
        Write-Host "  +======================================================+" -ForegroundColor $CN

        $c = Read-Host "`n  >> "
        switch ($c) {
            '1' { Reinstall-Klingo }
            '2' { Klingo-LimparCache }
            '3' { $global:KlingoPerfil = Get-Chrome-Profile-List } # Permite trocar se errar
            '0' { return }
        }
    } while ($true)
}


function Get-Chrome-Profile-List {
    $userDataPath = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    if (-not (Test-Path $userDataPath)) {
        Write-Host "    [!] Erro: Google Chrome não encontrado nesta máquina." -ForegroundColor $CR
        return $null
    }

    # Lista pastas que são perfis (Default ou Profile X)
    $perfis = Get-ChildItem -Path $userDataPath -Directory | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile*" }

    Write-SectionHeader "SELECIONE O PERFIL DO CHROME" $CC
    for ($i = 0; $i -lt $perfis.Count; $i++) {
        Write-Host "    [$($i+1)] $($perfis[$i].Name)" -ForegroundColor $CW
    }

    $escolha = Read-Host "`n    >> Numero do perfil onde o Klingo sera instalado"
    if ($escolha -match '^\d+$' -and [int]$escolha -le $perfis.Count) {
        return $perfis[[int]$escolha - 1].Name
    }
    return "Default" # Fallback para o padrão caso dê erro
}

function QuickFix-Menu {
    Show-Intro
    
    $Opcoes = @(
        " [1] Diagnostico de Hardware ",
        " [2] Status de Rede          ",
        " [3] Diagnostico Impressoras ",
        " [4] Otimizar RAM            ",
        " [5] Reparar Rede            ",
        " [6] Reparar Windows         ",
        " [7] Limpeza de Perfil       ",
        " [8] Klingo                  ",
        " [0] Sair                    "
    )
    
    $Selecao = 0

    while ($true) {
        [System.Console]::CursorVisible = $false
        Show-Header # Menu estático para não ter delay na seta
        
        Write-Host "  +======================================================+" -ForegroundColor $CP
        Write-Host "  |                  SISTEMA OPERACIONAL                 |" -ForegroundColor $CP
        Write-Host "  +------------------------------------------------------+" -ForegroundColor $CP
        
        for ($i = 0; $i -lt $Opcoes.Count; $i++) {
            if ($i -eq $Selecao) {
                # Item selecionado
                Write-Host "  > $($Opcoes[$i])" -BackgroundColor $CC -ForegroundColor Black
            } else {
                # Itens normais
                Write-Host "    $($Opcoes[$i])" -ForegroundColor $CW
            }
        }
        
        Write-Host "  +======================================================+" -ForegroundColor $CP

        $Tecla = [System.Console]::ReadKey($true).Key
        if ($Tecla -eq 'UpArrow')   { $Selecao-- }
        if ($Tecla -eq 'DownArrow') { $Selecao++ }
        if ($Selecao -lt 0) { $Selecao = $Opcoes.Count - 1 }
        if ($Selecao -ge $Opcoes.Count) { $Selecao = 0 }

        if ($Tecla -eq 'Enter') {
            [System.Console]::CursorVisible = $true 
            
            # Chamada ao Show-Glitch removida para maior agilidade

            switch ($Selecao) {
                0 { Show-SystemInfo }
                1 { Show-NetworkInfo }
                2 { Printer-Diagnostics-Menu }
                3 { Optimize-Memory }
                4 { Repair-Network }
                5 { Repair-Windows }
                6 { Clean-Profile }
                7 { Klingo-Menu }
                8 {
                    8 
                    # OPÇÃO [0] SAIR - FINALIZAÇÃO COMPLETA
                    Close-Report
                    Clear-Host
                    
                    Write-Host "`n  +======================================================+" -ForegroundColor $CP
                    Write-Host "  |                ENCERRANDO O QUICKFIX                 |" -ForegroundColor $CP
                    Write-Host "  +======================================================+" -ForegroundColor $CP
                    
                    Write-Host "`n    Encerrando o QuickFix, obrigado, " -NoNewline -ForegroundColor $CW
                    Write-Host "$($global:NomeTecnico)!" -ForegroundColor $CN
                    
                    if ($global:ArquivoSessao) {
                        Write-Host "`n    [i] O relatorio desta sessao foi gerado em:" -ForegroundColor $CC
                        Write-Host "    $($global:ArquivoSessao)" -ForegroundColor $CY
                    }
                    
                    Write-Host "`n    Sessao finalizada. Status: " -NoNewline -ForegroundColor $CG
                    Write-Host "[OFFLINE]" -ForegroundColor $CR
                    
                    Write-Host "  +======================================================+" -ForegroundColor $CP
                    
                    # Tempo maior (4s) para você conseguir ler o caminho do arquivo
                    Start-Sleep -Seconds 4 
                    exit 
                }
                }
            }
        }
    }

# Inicia o Sistema
QuickFix-Menu
