# 🛠 QuickFix

### Corporate Support & Diagnostic Script

**coded by emanoel peres :)**

![PowerShell](https://img.shields.io/badge/PowerShell-5.0+-5391FE?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%2B-0078D6?logo=windows&logoColor=white)
![Status](https://img.shields.io/badge/Status-Production-success)
![Scope](https://img.shields.io/badge/Scope-Internal-blue)
![License](https://img.shields.io/badge/License-MIT-red)
![Purpose](https://img.shields.io/badge/Purpose-Educational-orange)

---

## 📌 Sobre o Projeto

**QuickFix** é uma ferramenta PowerShell desenvolvida com fins **educacionais** durante estágio em T.I., com o objetivo de padronizar e agilizar procedimentos de suporte técnico em ambientes corporativos Windows.

O script centraliza diagnósticos, reparos e automações que normalmente exigiriam múltiplas ferramentas ou intervenções manuais, reduzindo o tempo de resolução de chamados e garantindo rastreabilidade completa das ações realizadas.

> ⚠️ **Aviso:** Este projeto foi desenvolvido para fins educacionais e de aprendizado. Nenhuma informação sensível de clientes ou organizações está presente no código.

---

## ✨ Funcionalidades

### 1️⃣ Diagnóstico de Hardware
- Sistema operacional, versão e build
- Processador, núcleos e placa-mãe
- Módulos de RAM por slot com velocidade
- Placa de vídeo e versão de driver
- Armazenamento com detecção SSD/HDD e alertas de espaço

### 2️⃣ Status de Rede
- Listagem de adaptadores ativos com IP e DNS
- Teste de gateway, conectividade e resolução DNS
- Detecção de DHCP estático ou dinâmico

### 3️⃣ Diagnóstico de Impressoras
- Status e fila de jobs por impressora
- Mapeamento de portas e IPs
- Teste de comunicação via ping
- Reinício do Spooler de impressão
- Limpeza de fila travada
- Definir impressora padrão
- Página de teste personalizada
- Forçar impressora online

### 4️⃣ Otimização de RAM
- Limpeza geral via Mem-Reduct Engine
- Limpeza cirúrgica de Chrome/Edge sem fechar o navegador
- Aplicar flags de economia de memória no Chrome

### 5️⃣ Reparo de Rede
- Reparo completo (Release/Renew/DNS/Winsock)
- Limpeza isolada de cache DNS
- Teste de conectividade detalhado
- Reset de adaptador de rede

### 6️⃣ Reparo do Windows
- DISM RestoreHealth com exit code real
- SFC /Scannow com interpretação de resultado
- Limpeza de componentes WinSxS
- Resultados reais em vez de mensagens fixas

### 7️⃣ Limpeza de Perfil
- Limpeza de `%TEMP%` do usuário
- Limpeza de `C:\Windows\Temp`
- Limpeza de Prefetch
- Limpeza completa em sequência
- Cálculo de espaço antes/depois em cada operação

### 8️⃣ Klingo
- Reinstalação limpa do Klingo como PWA no Chrome
- Desinstalação via registro com App ID correto
- Abertura direta no perfil correto sem tela de seleção
- Guia visual passo a passo para o técnico
- Limpeza de cache, Local Storage, IndexedDB e Service Worker

---

## 📋 Sistema de Relatórios por Sessão

Cada sessão gera **um único relatório** com todas as ações realizadas:

```
C:\services\relatorios\
    NomeDoTecnico\
        sessao_27022026_143022.txt
        sessao_25022026_090511.txt
```

O relatório registra:
- Data e hora de início e encerramento da sessão
- Nome do técnico, máquina e IP
- Cada ação realizada com horário e resultado real

---

## ▶ Execução

### Método 1 — Direto na memória (proxy-friendly)
```powershell
$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
IEX $wc.DownloadString('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1')
```

### Método 2 — Download e execução local (auditável)
```powershell
$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
$wc.DownloadFile('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1', 'QuickFIX.ps1')
powershell -ExecutionPolicy Bypass -File .\QuickFIX.ps1
```

### Método 3 — Via Invoke-RestMethod
```powershell
powershell -ExecutionPolicy Bypass -Command "iex (irm https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1 -Proxy $null)"
```

### Método 4 — Com verificação SHA256 (enterprise-safe)
```powershell
$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
$wc.DownloadFile('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1', 'QuickFIX.ps1')
$hash = Get-FileHash .\QuickFIX.ps1 -Algorithm SHA256
if ($hash.Hash -eq 'INSERT_OFFICIAL_SHA256') {
    powershell -ExecutionPolicy Bypass -File .\QuickFIX.ps1
} else {
    Write-Host "ERROR: File hash mismatch! Aborting execution."
}
```

---

## ⚙ Requisitos

| Requisito | Versão mínima |
|---|---|
| Windows | 10 ou superior |
| PowerShell | 5.0 ou superior |
| Privilégios | Administrador local |
| ExecutionPolicy | Bypass |

---

## 🔧 Configuração do Klingo

Para adaptar o módulo 8 a outro ambiente, edite as variáveis no topo da seção do Klingo:

```powershell
$KlingoURL    = "https://sua-url.klingo.app/#/"
$KlingoPerfil = "Profile 3"   # Perfil do Chrome onde o Klingo está instalado
$KlingoAppId  = "SEU_APP_ID"  # Obtido via registro do Windows
$KlingoNome   = "klingo"      # Nome como aparece em chrome://apps
```

**Como descobrir o App ID e o perfil:**
```powershell
Get-ChildItem "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall" |
ForEach-Object { Get-ItemProperty $_.PSPath } |
Where-Object { $_.DisplayName -like "*klingo*" } |
Select-Object -ExpandProperty UninstallString
```

---

## 🏗 Arquitetura

QuickFix é composto por **8 módulos independentes** acessíveis via menu numérico. Cada módulo executa em isolamento para reduzir risco operacional.

```
QuickFix-Menu
├── Show-Intro           # Identificação do técnico + Init-Report
├── Init-Report          # Inicializa sessão de log
├── Send-Report          # Anexa ação ao log da sessão
├── Close-Report         # Fecha sessão com rodapé
├── Show-SystemInfo      # Módulo 1
├── Show-NetworkInfo     # Módulo 2
├── Printer-Diagnostics-Menu  # Módulo 3
├── Optimize-Memory      # Módulo 4
├── Repair-Network       # Módulo 5
├── Repair-Windows       # Módulo 6
├── Clean-Profile        # Módulo 7
└── Klingo-Menu          # Módulo 8
    ├── Reinstall-Klingo
    └── Klingo-LimparCache
```

---

## 📌 Princípios do Projeto

- **Padronização operacional** — mesmo procedimento em qualquer máquina
- **Mínima intervenção** — nenhuma ação impactante sem confirmação do técnico
- **Rastreabilidade total** — log completo de sessão por técnico
- **Resultados reais** — sem mensagens fixas de sucesso, sempre o resultado real
- **Modularidade** — cada módulo é independente e isolado

---

## 📝 Notas

- Nenhuma informação sensível de organizações está presente no código
- As configurações específicas de ambiente (URLs, perfis, IDs) devem ser adaptadas localmente
- Desenvolvido como projeto educacional de estágio em T.I.
- Restrito a uso interno pelo departamento de T.I.

---

## About

Modular PowerShell toolkit for Windows 10/11 environments focused on endpoint diagnostics, network recovery, service management, logging, and safe remediation workflows.
