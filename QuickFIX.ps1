# ============================
# QuickFIX - Diagnóstico e Ferramentas
# ============================

# Função de pausa
function Pause {
    Write-Host "`nPressione qualquer tecla para continuar..." -NoNewline
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# Função de Log
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    $logFile = "$PSScriptRoot\QuickFix.log"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Type] - $Message" | Out-File -Append -FilePath $logFile
}

# Função para consulta segura ao CIM (WMI)
function Safe-CimQuery {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ClassName
    )
    $resultado = @()
    try {
        $resultado = Get-CimInstance -ClassName $ClassName
    } catch {
        # Captura a exceção e acessa diretamente o erro
        $errorMessage = $_.Exception.Message
        $logMessage = "Erro ao consultar $($ClassName): $errorMessage"
        Write-Log $logMessage "ERROR"
        Write-Host "$($errorMessage)" -ForegroundColor Red
    }
    return $resultado
}

# Função para criar ponto de restauração
function Create-RestorePoint {
    # Cria um ponto de restauração (mockup para o QuickFIX)
    Write-Host "`nPonto de restauração criado com sucesso (Mockup)." -ForegroundColor Green
    Write-Log "Restore Point criado"
    Pause
}

# Função para exibir informações do sistema
function Show-SystemInfo {
    Clear-Host
    Write-Host "`n=== Informações do Sistema ===" -ForegroundColor Cyan

    # CPU
    Write-Host "`n[Hardware]" -ForegroundColor Green
    $cpu = Safe-CimQuery "Win32_Processor" | Select-Object Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, AddressWidth, L2CacheSize, L3CacheSize, CacheSize
    if ($cpu) {
        Write-Host "CPU:" 
        Write-Host "  Nome:          $($cpu.Name)"
        Write-Host "  Fabricante:    $($cpu.Manufacturer)"
        Write-Host "  Núcleos físicos: $($cpu.NumberOfCores)"
        Write-Host "  Threads:       $($cpu.NumberOfLogicalProcessors)"
        Write-Host "  Velocidade:    $($cpu.MaxClockSpeed) MHz"
        Write-Host "  Arquitetura:   $($cpu.AddressWidth)-bit"
        Write-Host "  Cache L2:      $($cpu.L2CacheSize) KB"
        Write-Host "  Cache L3:      $($cpu.L3CacheSize) KB"
        Write-Host "  Cache Total:   $($cpu.CacheSize) KB"
    }

    # RAM
    $cs = Safe-CimQuery "Win32_ComputerSystem"
    if ($cs) {
        $ramBytes = $cs.TotalPhysicalMemory
        $ramGB = [math]::Round($ramBytes / 1GB, 2)
        $ramModules = Safe-CimQuery "Win32_PhysicalMemory" | Measure-Object | Select-Object -ExpandProperty Count
        $ramDetails = Safe-CimQuery "Win32_PhysicalMemory"
        Write-Host "RAM Total:      $ramGB GB"
        Write-Host "Módulos físicos: $ramModules"
        foreach ($module in $ramDetails) {
            Write-Host "`nMódulo de RAM: $($module.DeviceLocator)"
            Write-Host "  Tamanho:      $([math]::Round($module.Capacity / 1GB, 2)) GB"
            Write-Host "  Tipo:         $($module.MemoryType)"
            Write-Host "  Velocidade:   $($module.Speed) MHz"
        }
    }

    # Placa-mãe / BIOS
    Write-Host "[Placa-mãe / BIOS]" -ForegroundColor Green
    $board = Safe-CimQuery "Win32_BaseBoard"
    $biosRaw = Safe-CimQuery "Win32_BIOS"
    $biosDate = "Não disponível"
    if ($biosRaw -and $biosRaw.ReleaseDate) {
        try { $biosDate = ([Management.ManagementDateTimeConverter]::ToDateTime($biosRaw.ReleaseDate)).ToString("dd/MM/yyyy") } catch {}
    }
    Write-Host "Placa-mãe:      $($board.Product)"
    Write-Host "Fabricante:     $($board.Manufacturer)"
    Write-Host "Número de série: $($board.SerialNumber)"
    Write-Host "BIOS:           $($biosRaw.SMBIOSBIOSVersion)"
    Write-Host "Data BIOS:      $biosDate`n"

    # Armazenamento - Discos e Partições
    Write-Host "[Armazenamento]" -ForegroundColor Green
    $disks = Safe-CimQuery "Win32_DiskDrive"
    foreach ($disk in $disks) {
        Write-Host "`nDisco:          $($disk.Model)"
        Write-Host "  Tipo:          $($disk.MediaType)"
        Write-Host "  Capacidade:    $([math]::Round($disk.Size / 1GB, 2)) GB"
        $partitions = Safe-CimQuery "Win32_DiskPartition" | Where-Object {$_.DiskIndex -eq $disk.Index}
        foreach ($partition in $partitions) {
            Write-Host "  Partição:      $($partition.Name)"
            Write-Host "    Tamanho:     $([math]::Round($partition.Size / 1GB, 2)) GB"
            Write-Host "    Sistema de Arquivos: $($partition.FileSystem)"
        }
    }

    # GPU - Placa de Vídeo
    Write-Host "`n[Placa de Vídeo]" -ForegroundColor Green
    $gpus = Safe-CimQuery "Win32_VideoController"
    foreach ($gpu in $gpus) {
        Write-Host "  Nome:           $($gpu.Name)"
        Write-Host "  Fabricante:     $($gpu.AdapterCompatibility)"
        Write-Host "  Memória:        $([math]::Round($gpu.AdapterRAM / 1MB)) MB"
        Write-Host "  Driver:         $($gpu.DriverVersion)"
        Write-Host "  Status:         $($gpu.Status)`n"
    }

    Pause
}

# Função para exibir informações de rede (Ethernet)
function Show-NetworkInfo {
    Clear-Host
    Write-Host "`n=== Informações de Conexão Ethernet ===" -ForegroundColor Cyan

    # Obtém os adaptadores de rede habilitados
    $adapters = Safe-CimQuery "Win32_NetworkAdapterConfiguration" | Where-Object { $_.IPEnabled -eq $true }
    
    if (-not $adapters) {
        Write-Host "Nenhuma interface ativa encontrada." -ForegroundColor Yellow
        Write-Log "Usuário visualizou informações de rede: nenhuma interface ativa."
        Pause
        return
    }

    # Exibindo informações sobre cada adaptador de rede
    foreach ($adapter in $adapters) {
        Write-Host "Adaptador:      $($adapter.Description)"
        Write-Host "Status:         $($adapter.Status)"
        Write-Host "MAC:            $($adapter.MACAddress)"
        Write-Host "Gateway:        $($adapter.DefaultIPGateway -join ', ')"
        Write-Host "DHCP Habilitado: $($adapter.DHCPEnabled)"
        Write-Host "DNS:            $($adapter.DNSServerSearchOrder -join ', ')"

        # Exibe endereços IP configurados (IPv4, IPv6)
        if ($adapter.IPAddress) {
            Write-Host "`nEndereços IP configurados: " -ForegroundColor Cyan
            foreach ($ip in $adapter.IPAddress) {
                Write-Host "  - $ip" -ForegroundColor Yellow
            }
        }

        # Exibe a máscara de sub-rede
        if ($adapter.IPSubnet) {
            Write-Host "`nMáscaras de Sub-rede:" -ForegroundColor Cyan
            foreach ($subnet in $adapter.IPSubnet) {
                Write-Host "  - $subnet" -ForegroundColor Yellow
            }
        }

        # Teste de conectividade - Ping para o Gateway
        $gateway = $adapter.DefaultIPGateway[0]  # Pega o primeiro gateway
        if ($gateway) {
            Write-Host "`nTestando conectividade com o gateway ($gateway)..." -ForegroundColor Cyan
            try {
                $pingResult = Test-Connection -ComputerName $gateway -Count 2 -Quiet
                if ($pingResult) {
                    Write-Host "Conexão com o gateway $gateway bem-sucedida!" -ForegroundColor Green
                } else {
                    Write-Host "Falha na conexão com o gateway $gateway." -ForegroundColor Red
                }
            } catch {
                Write-Host "Erro ao testar a conectividade com o gateway $gateway." -ForegroundColor Red
            }
        }

        Write-Host "`n"  # Linha em branco entre adaptadores
    }

    Write-Log "Usuário visualizou informações de rede."
    Pause
}

# Função para a tela de entrada e rollback
function Rollback-Welcome {
    Clear-Host
    Write-Host "Olá, bem-vindo ao QuickFIX.`n" -ForegroundColor Cyan
    Write-Host "Recomendamos criar um ponto de restauração antes de alterar o sistema.`n" -ForegroundColor Yellow
    Write-Host "[1] - Criar Restore Point"
    Write-Host "[2] - Pular (Não recomendado)`n"

    $choice = Read-Host "Escolha uma opção"
    if ($choice -notmatch '^[12]$') { 
        Write-Host "Opção inválida!"; 
        Pause 
        Rollback-Welcome  # Chama novamente caso a opção seja inválida
        return
    }

    switch ($choice) {
        '1' {
            Create-RestorePoint
            break
        }
        '2' {
            Write-Host "`nRollback pulado." -ForegroundColor Yellow
            Write-Log "Restore Point pulado"
            break
        }
    }

    # Após o ponto de restauração, vai para o menu principal
    QuickFix-Menu
}

# Função Menu principal
function QuickFix-Menu {
    do {
        Clear-Host
        Write-Host "=== QuickFix Tools ===" -ForegroundColor Cyan
        Write-Host "[1] - Mostrar Informações do Sistema"
        Write-Host "[2] - Mostrar Informações de Conexão Ethernet"
        Write-Host "[0] - Sair" -ForegroundColor Cyan

        $choice = Read-Host "Escolha uma opção"
        if ($choice -notmatch '^\d+$') { Write-Host "Digite apenas números válidos!"; Pause; continue }

        switch ($choice) {
            '1' { Show-SystemInfo }
            '2' { Show-NetworkInfo }
            '0' { Write-Host "`nSaindo do QuickFIX..." -ForegroundColor Cyan; break }
            default { Write-Host "Opção inválida"; Pause }
        }
    } while ($true)
}

# INÍCIO
Rollback-Welcome
