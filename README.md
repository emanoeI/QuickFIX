<h1 align="center">QuickFix</h1>
<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=for-the-badge&logo=powershell" alt="PowerShell Version">
  <img src="https://img.shields.io/badge/Environment-Windows-0078D6?style=for-the-badge&logo=windows" alt="OS Windows">
  <img src="https://img.shields.io/badge/Project-Educational-green?style=for-the-badge" alt="Educational Project">
</p>
<p align="center">
  <strong>Solução Script para Suporte Técnico e Automação de Infraestrutura</strong><br>
  <i>Desenvolvido como projeto de conclusão de estágio em T.I. para otimização de fluxos operacionais.</i>
</p>

---

## 📺 QuickFix em ação

<p align="center">
  <img src="./Animação.gif" alt="Demonstração do QuickFix" width="700px">
</p>

---

## 📖 Sobre o Projeto

O **QuickFix** é uma ferramenta de terminal interativa desenvolvida para agilizar o atendimento técnico N1 e N2. O projeto foca em consolidar diagnósticos complexos e rotinas de reparo em uma interface unificada, reduzindo o tempo de resposta em ambientes críticos como farmácias e clínicas.

> **Foco Educacional:** Este projeto explora o uso avançado de **PowerShell scripting**, consultas **CIM/WMI**, manipulação de processos via **.NET Interop** e sistemas de auditoria por logs.

---

## 🚀 Módulos Implementados

| Categoria | Funcionalidades |
| :--- | :--- |
| 🖥️ **Hardware** | Mapeamento completo de CPU, RAM, GPU e integridade de armazenamento (SSD/HDD). |
| 🌐 **Rede** | Reset de stack TCP/IP, limpeza de cache DNS, renovação de concessão DHCP e testes de download e upload com motor Ookla (Speedtest CLI). |
| 🖨️ **Impressoras** | Diagnóstico de filas, mapeamento de portas IP, reinício de Spooler e envio de página de teste exclusiva para suporte. |
| 🧠 **Otimização** | Motor de limpeza de memória RAM e aplicação de flags de performance para navegadores. |
| 🛠️ **Reparo OS** | Automatização de rotinas DISM e SFC com interpretação de códigos de saída. |
| ☁️ **Klingo PWA** | Gestão inteligente de cache e reinstalação dinâmica baseada em perfis do Chrome. |

---

## 🔐 Segurança e Auditoria

* **Acesso Autenticado:** Proteção via senha invisível (`-AsSecureString`) garantindo que apenas técnicos autorizados utilizem as ferramentas de reparo.
* **Execução em Memória:** O script é carregado diretamente do GitHub e executado via `iex`, sem salvar código-fonte no disco da máquina.
* **Log de Sessão:** Cada ação executada é documentada em um relatório `.txt` individual, incluindo data, hora e o nome do operador, salvo em `C:\services\relatorios\`.

---

## 📥 Como Utilizar

### Pré-requisitos

* **Sistema:** Windows 10 ou 11.
* **PowerShell:** Versão 5.1 ou superior (já incluso no Windows).
* **Privilégios:** Administrador (solicitado automaticamente via UAC).

---

### ▶️ Método 1 — Execução Direta (Recomendado)

Cole o comando abaixo no PowerShell e pressione Enter:

```powershell
irm "https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1" | iex
```

> Nenhum arquivo é salvo no disco. O script executa inteiramente em memória.

---

### 🖱️ Método 2 — Atalho `.bat` (Para Técnicos)

Crie um arquivo `QuickFix.bat` com o conteúdo abaixo e execute com duplo clique:

```batch
@echo off
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1' | iex"
```

> O UAC será solicitado automaticamente caso não esteja em modo administrador.

---

### 💾 Método 3 — Download Manual

1. Baixe o arquivo [`QuickFIX.ps1`](https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1).
2. Clique com o botão direito → **"Executar com o PowerShell"**.
3. Confirme a elevação de privilégios quando solicitado.

---

### 🔑 Primeiro Acesso

Ao iniciar, informe sua **identificação** e a **chave de acesso** fornecida pelo administrador do sistema.

---

## 🛠️ Tecnologias e Conceitos Aplicados

* **ANSI Escaping:** Interface visual com paleta de cores personalizada e navegação por teclado.
* **WMI/CIM:** Coleta detalhada de informações do sistema operacional e hardware.
* **Ookla Speedtest CLI:** Medição precisa de download, upload e latência via motor oficial da Speedtest.
* **SecureString / .NET Interop:** Manipulação segura de credenciais em memória sem exposição em plaintext.
* **Error Handling:** Blocos `try/catch/finally` para garantir resiliência em ambientes críticos.
* **Execução Remota:** Carregamento via `irm | iex` direto do GitHub, sem instalação ou arquivos locais.

---

## 👨‍💻 Autor

**Emanoel Peres**


---

<p align="center">
  <i>"Transformando processos complexos em comandos de um clique."</i>
</p>
